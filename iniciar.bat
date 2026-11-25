@echo off
cd /d "%~dp0"
echo ==== Paso 1: comprobando Python ====
py --version || python --version
echo.
echo ==== Paso 2: instalando dependencias ====
py -m pip install -r requirements.txt || python -m pip install -r requirements.txt
echo.
echo ==== Paso 3: iniciando la tienda ====
echo NO CIERRES esta ventana mientras uses la tienda.
echo Cuando veas "Running on http://127.0.0.1:5000" ya puedes entrar.
echo.
start "" cmd /c "timeout /t 4 >nul & start http://127.0.0.1:5000"
py app.py || python app.py
echo.
echo *** La tienda se detuvo. Si ves un error arriba, copialo y envialo. ***
pause
