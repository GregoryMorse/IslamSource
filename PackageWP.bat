CD ..
DEL IslamSource\IslamSourceWP.zip
REM No spaces in paths for non-Windows compatibility
MKDIR IslamSource\IslamResources\MyProject
COPY "IslamSource\IslamResources\My Project\Resources.resx" IslamSource\IslamResources\MyProject\Resources.resx
"%ProgramFiles%\7-Zip\7z.exe" a IslamSource\IslamSourceWP.zip IslamSource\cache IslamSource\readme.txt IslamSource\IslamSource.jpg IslamSource\IslamSource2.jpg IslamSource\*.php IslamSource\isbutton.js "IslamSource\IslamResources\MyProject\Resources.resx" IslamSource\HostPage\metadata\ArabicData.xml IslamSource\metadata\islaminfo.xml IslamSource\metadata\quran-hafs.xml IslamSource\IslamResources\*.resx IslamSource\files\*.ttf IslamSource\files\*.otf
DEL IslamSource\IslamResources\MyProject\Resources.resx
RMDIR IslamSource\IslamResources\MyProject
CD IslamSource