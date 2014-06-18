cd bin
e:\wsusoffline\bin\wget.exe -N -nd http://tanzil.net/res/text/metadata/quran-data.xml
e:\wsusoffline\bin\wget.exe -N -nd --content-disposition -Oquran-uthmani.xml --post-data=quranType=uthmani^&marks=^&sajdah=^&alef=true^&me_quran=^&outType=xml^&agree=true http://tanzil.net/pub/download/download.php
e:\wsusoffline\bin\wget.exe -N -nd http://tanzil.net/trans/
SETLOCAL ENABLEDELAYEDEXPANSION
FOR /F delims^=^"^ tokens^=6^,12 %%A IN ('FIND "title=" index.html') DO ((IF "%%B"=="return download(this)" (SET VAR="%%A") ELSE (SET VAR="%%B")) & e:\wsusoffline\bin\wget.exe -N -nd --content-disposition http://tanzil.net/trans/?transID=!VAR:~25!^&type=txt)
ENDLOCAL
del index.html*
cd ..
cd tanzilxml
e:\wsusoffline\bin\wget.exe -N -nd http://tanzil.net/trans/
SETLOCAL ENABLEDELAYEDEXPANSION
FOR /F delims^=^"^ tokens^=6^,12 %%A IN ('FIND "title=" index.html') DO ((IF "%%B"=="return download(this)" (SET VAR="%%A") ELSE (SET VAR="%%B")) & e:\wsusoffline\bin\wget.exe -N -nd --content-disposition http://tanzil.net/trans/?transID=!VAR:~25!^&type=xml)
ENDLOCAL
del index.html*
cd ..