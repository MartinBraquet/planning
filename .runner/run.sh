#!/usr/bin/env bash
set -Eeuo pipefail

cd $(dirname $0)
pwd

echo "Génération des horaires..."
echo

# ensure folders, error if no planning folder
if [[ ! -d "../planning" ]]; then
  echo "ERREUR: Le dossier 'planning' est introuvable."
  echo "Veuillez créer le dossier 'planning' et y placer vos fichiers .xlsx."
  exit 1
fi


# Create venv if missing
if [[ ! -d ".venv" ]]; then
  echo "Creating virtual environment..."
  python3 -m venv .venv || {
    echo "ERROR: Failed to create virtual environment."
    exit 1
  }

  # activate
  source .venv/bin/activate

  echo "Upgrading pip..."
  python -m pip install --upgrade pip

  echo "Installing required packages..."
  if [[ -f "requirements.txt" ]]; then
    pip install -r requirements.txt
  elif [[ -f "../requirements.txt" ]]; then
    pip install -r ../requirements.txt
  else
    echo "WARNING: requirements.txt not found; skipping dependency install."
  fi
else
  # activate existing venv
  source .venv/bin/activate
fi

# Show Python version
python --version

# Run your script
python run.py

echo
echo 'Fini. Voir fichiers dans le dossier "planning".'