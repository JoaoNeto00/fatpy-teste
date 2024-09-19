@echo off
echo Instalando as bibliotecas Python necessarias...

REM Instalação do pandas
pip install --trusted-host=pypi.org --trusted-host=files.pythonhosted.org --user pandas
echo -----------concluido pandas-----------

REM Instalação do tkinter
pip install --trusted-host=pypi.org --trusted-host=files.pythonhosted.org --user tk
echo -----------concluido tkinter-----------

REM Instalação do openpyxl
pip install --trusted-host=pypi.org --trusted-host=files.pythonhosted.org --user openpyxl
echo -----------concluido openpyxl-----------

REM Instalação pillow
pip install --trusted-host=pypi.org --trusted-host=files.pythonhosted.org --user pillow
echo -----------concluido pillow-----------

REM Instalação pythony
pip install --trusted-host=pypi.org --trusted-host=files.pythonhosted.org --user pythonw
echo -----------concluido pythonw-----------

REM Instalação ttkbootstrap
pip install --trusted-host=pypi.org --trusted-host=files.pythonhosted.org --user ttkbootstrap
echo -----------concluido ttkbootstrap-----------

echo Instalacao concluida.
pause
