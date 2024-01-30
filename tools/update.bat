@ECHO OFF

SET ARCHIVE=\\server\movies\library\eHome.zip
SET FOLDER=\\192.168.0.201\C$\Users

DIR "%ARCHIVE%"
ECHO.

PUSHD "%FOLDER%"
REM CD
FOR /D %%i in (*) DO (ENDLOCAL
  IF NOT "%%i" == "Public" (CALL :UPDATE %%i) ELSE (ECHO "Skipping %%i")
)
POPD
PAUSE
EXIT /b

:UPDATE
PUSHD "%*\AppData\Roaming\Microsoft\eHome"
CD
FOR %%i in (DvdCoverCache, DvdInfoCache) DO ( RMDIR /S /Q "%%i" )
TAR -xf "%ARCHIVE%"
POPD
EXIT /b
