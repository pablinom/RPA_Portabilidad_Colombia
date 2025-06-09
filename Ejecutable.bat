REM

set original_dir=%CD%
set venv_root_dir="C:\RPA_PORTACOL_INTERACTIVO\env"

cd %venv_root_dir%

call %venv_root_dir%\Scripts\activate.bat

python C:\RPA_PORTACOL_INTERACTIVO\src\RPA_PORTACOL\main.py

cd %original_dir%

exit /B 1
