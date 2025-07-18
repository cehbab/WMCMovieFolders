@ECHO OFF

SET ARCHIVE=\\server\movies\library\eHome.zip
SET FOLDER=\\192.168.0.201\C$\Users
REM SET FOLDER=C:\Users

DIR "%ARCHIVE%"
ECHO.

PUSHD "%FOLDER%"
IF NOT "%1" == "" (
  IF EXIST "%1" (CALL :UPDATE "%1") ELSE (
    ECHO Missing %1
    ECHO Available:
    FOR /D %%i in (*) DO (ENDLOCAL
      IF NOT "%%i" == "Public" (ECHO %%i)
    )
  )
) ELSE (
  FOR /D %%i in (*) DO (ENDLOCAL
    SET SKIP=
    IF "%%i" == "Administrator" SET SKIP=1
    IF "%%i" == "Public" SET SKIP=1
    IF "%%i" == "nx" SET SKIP=1
    IF DEFINED SKIP (ECHO Skipping %%i) ELSE (CALL :UPDATE %%i)
  )
)
POPD
PAUSE
EXIT /b


:UPDATE
PUSHD "%1\AppData\Roaming\Microsoft\eHome"
CD
FOR %%i in (DvdCoverCache, DvdInfoCache) DO ( RMDIR /S /Q "%%i" )
TAR -xf "%ARCHIVE%"
POPD
EXIT /b
