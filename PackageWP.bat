CD ..
DEL IslamSource\IslamSourceWP.zip
REM No spaces in paths for non-Windows compatibility
MKDIR IslamSource\IslamResources\MyProject
COPY "IslamSource\IslamResources\My Project\Resources.resx" IslamSource\IslamResources\MyProject\Resources.resx
"%ProgramFiles%\7-Zip\7z.exe" a IslamSource\IslamSourceWP.zip IslamSource\cache IslamSource\readme.txt IslamSource\*.php IslamSource\isbutton.js "IslamSource\IslamResources\MyProject\Resources.resx" IslamSource\metadata\islaminfo.xml IslamSource\metadata\quran-hafs.xml IslamSource\metadata\quran-data.xml IslamSource\metadata\*.txt IslamSource\IslamResources\*.resx IslamSource\files\*.ttf IslamSource\files\*.otf
DEL IslamSource\IslamResources\MyProject\Resources.resx
RMDIR IslamSource\IslamResources\MyProject
CD IslamSource

IF NOT "%1"=="" (
svn checkout http://plugins.svn.wordpress.org/islamsource/trunk/
"%ProgramFiles%\7-Zip\7z.exe" x -otrunk IslamSourceWP.zip
move /Y trunk\IslamSource\*.* trunk
for /D %%F in (trunk\IslamSource\*) do if EXIST trunk\%%~nxF (move /Y %%F\*.* trunk\%%~nxF && rmdir /s /q %%F) else (move /Y %%F trunk)
rmdir /s /q trunk\IslamSource
svn add trunk\*
for /D %%F in (trunk\*) do svn add %%F\*
svn commit trunk -m "%1"
svn copy http://plugins.svn.wordpress.org/islamsource/trunk/ http://plugins.svn.wordpress.org/islamsource/tags/%2 -m "%1"
rmdir /s /q trunk
svn checkout http://plugins.svn.wordpress.org/islamsource/assets
copy *.jpg assets
svn add assets\*.jpg
svn commit assets -m "%1"
rmdir /s /q assets
)