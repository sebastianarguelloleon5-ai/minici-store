#!/bin/bash
# Copia diaria de los datos (base de datos + fotos). Se programa con cron: ver la guía.
set -e
APP="$HOME/minici-store"
DEST="$HOME/respaldos"
FECHA=$(date +%F)
mkdir -p "$DEST"
sqlite3 "$APP/minici_store.db" ".backup '$DEST/minici_$FECHA.db'"
tar -czf "$DEST/fotos_$FECHA.tar.gz" -C "$APP" fotos_productos 2>/dev/null || true
find "$DEST" -type f -mtime +14 -delete
echo "Respaldo del $FECHA guardado en $DEST"
