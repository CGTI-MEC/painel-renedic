@echo off
chcp 65001 >nul
:: ============================================================
:: PAINEL RENEDIC — ATUALIZAR DADOS E PUBLICAR NO GITHUB
:: Duplo clique para executar
:: ============================================================

cd /d "%~dp0"

echo.
echo ============================================================
echo  PAINEL RENEDIC — Passo 1/2: processando os dados
echo ============================================================
echo.

:: Descobre qual comando do Python existe neste computador
set PY=
py -3.11 --version >nul 2>&1 && set "PY=py -3.11"
if not defined PY py --version >nul 2>&1 && set "PY=py"
if not defined PY python --version >nul 2>&1 && set "PY=python"
if not defined PY (
    echo ERRO: Python nao encontrado neste computador.
    pause
    exit /b 1
)

%PY% processar.py
if errorlevel 1 (
    echo.
    echo ERRO no processamento. Leia a mensagem acima. Nada foi publicado.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  Passo 2/2: enviando para o GitHub
echo ============================================================
echo.

git add saida/redes.json saida/envios.json saida/config.json saida/meta.json
git diff --cached --quiet
if not errorlevel 1 (
    echo Nenhuma mudanca nos dados desde a ultima publicacao. Nada a enviar.
    echo.
    pause
    exit /b 0
)

git commit -m "Atualizacao dos dados - %date% %time:~0,5%"
git push origin main
if errorlevel 1 (
    echo.
    echo ERRO ao enviar para o GitHub. Verifique a internet e o login no GitHub.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  CONCLUIDO! O painel estara atualizado em 1 a 2 minutos.
echo ============================================================
echo.
pause
