#!/usr/bin/env bash
set -e  # detener ejecución si ocurre un error

ENV_NAME="mi_entorno"
YAML_FILE="environment.yml"

echo "🔍 Verificando archivo YAML…"
if [[ ! -f "$YAML_FILE" ]]; then
    echo "❌ ERROR: No se encontró $YAML_FILE"
    exit 1
fi

echo "🧹 Eliminando environment previo (si existe)…"
mamba env remove -n "$ENV_NAME" --yes || true

echo "📦 Creando nuevo environment '$ENV_NAME' desde $YAML_FILE…"
mamba env create -n "$ENV_NAME" -f "$YAML_FILE"

echo "✅ Environment creado correctamente."
echo "➡️ Para activarlo:"
echo "   conda activate $ENV_NAME"
