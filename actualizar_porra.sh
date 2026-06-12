#!/bin/bash
# ============================================================
# actualizar_porra.sh
# Ejecuta el calculador de puntos y actualiza data.json
# Para programar diariamente con cron:
#   crontab -e
#   0 8 * * * /ruta/completa/a/actualizar_porra.sh
# ============================================================

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
LOG_FILE="$SCRIPT_DIR/log_actualizacion.txt"

echo "$(date '+%Y-%m-%d %H:%M:%S') - Iniciando actualización..." >> "$LOG_FILE"

cd "$SCRIPT_DIR"
python3 scripts/calcular_puntos.py >> "$LOG_FILE" 2>&1

if [ $? -eq 0 ]; then
    echo "$(date '+%Y-%m-%d %H:%M:%S') - ✓ Actualización completada OK" >> "$LOG_FILE"
else
    echo "$(date '+%Y-%m-%d %H:%M:%S') - ✗ ERROR en la actualización" >> "$LOG_FILE"
fi

echo "─────────────────────────────" >> "$LOG_FILE"
