#!/bin/bash

echo "Starting BidAnalysis Tool (Python 3.12 preferred)..."
echo

# Prefer Python 3.12 venv on Unix-like systems if available
if command -v python3.12 >/dev/null 2>&1; then
    if [ ! -d ".venv312" ]; then
        echo "Creating Python 3.12 virtual environment..."
        python3.12 -m venv .venv312
    fi
    # shellcheck disable=SC1091
    source .venv312/bin/activate
else
    # Fallback to default python3
    if [ ! -d ".venv" ]; then
        echo "Creating virtual environment with system python3..."
        python3 -m venv .venv
    fi
    # shellcheck disable=SC1091
    source .venv/bin/activate
fi

# Install dependencies
echo "Installing dependencies..."
pip install -r backend/requirements.txt

# Check if .env file exists
if [ ! -f "backend/.env" ]; then
    echo "Creating .env file from template..."
    cp "backend/.env.template" "backend/.env"
    echo
    echo "Please edit backend/.env and set API keys (DASHSCOPE_API_KEY or ARK_API_KEY)"
    echo "Then run this script again."
    exit 1
fi

# Start the application
echo "Starting backend server..."
cd backend || exit 1
python run.py
