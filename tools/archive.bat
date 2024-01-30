@ECHO OFF

SET ARCHIVE=\\server\movies\library\eHome.zip

PUSHD "%APPDATA%\Microsoft\eHome"
CD
TAR -caf "%ARCHIVE%" DvDCoverCache DvdInfoCache
POPD

PAUSE
EXIT /b
