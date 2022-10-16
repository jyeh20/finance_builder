@ECHO off

SETLOCAL ENABLEDELAYEDEXPANSION

:: get today's date
SET index=0
SET month=
SET months=January February March April May June July August September October November December
for %%m in (%months%) do (
  SET month_arr[!index!]=%%m
  SET /a index+=1
)

for /f "tokens=2,4 delims=/ " %%a in ("%date%") do (
  SET month=%%a
  SET year=%%b
)

SET /a month-=1
SET month=!month_arr[%month%]!
ECHO !month!
ECHO !year!

findstr "user_path" paths.config > path.txt

SET /P user_path=<path.txt
SET user_path=%user_path:~10%
DEL path.txt

SET PATH="%user_path%\Documents\Personal\%year%\Finances\%month%.xlsx"
SET FINANCE_DIR="%user_path%\Documents\Personal\%year%\Finances"
SET FILE_TO_OPEN="%month%.xlsx"

if exist %PATH% (
  ECHO "Directory exists, opening..."
  GOTO OPEN
) else (
  ECHO "Directory does not exist"
  GOTO COPY_TEMPLATE
)

:COPY_TEMPLATE
ECHO "Copying template..."
copy "%user_path%\Documents\Personal\Finances_Template.xlsx" %FINANCE_DIR%
REN %FINANCE_DIR%\Finances_Template.xlsx %FILE_TO_OPEN%

:OPEN
ECHO "Opening %PATH%"
CD %FINANCE_DIR%
START excel.exe %FILE_TO_OPEN%
GOTO END

:END