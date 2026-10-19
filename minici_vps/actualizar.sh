#!/bin/bash
# Actualiza Minici Store desde GitHub y reinicia la app.  Uso:  ./actualizar.sh
set -e
cd "$(dirname "$0")"
echo ">> Descargando cambios de GitHub..."
git pull --ff-only
echo ">> Instalando dependencias..."
.venv/bin/pip install -q -r requirements.txt
echo ">> Reiniciando la app..."
sudo systemctl restart minici
sleep 3
sudo systemctl is-active minici && echo ">> Listo: la app está corriendo."
