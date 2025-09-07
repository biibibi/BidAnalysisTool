#!/bin/bash

echo "Starting BidAnalysis Tool..."
echo

# Check if virtual environment exists
if [ ! -d "venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv venv
fi

# Activate virtual environment
source venv/bin/activate

# Install dependencies
echo "Installing dependencies..."
pip install -r backend/requirements.txt

# Check if .env file exists
if [ ! -f "backend/.env" ]; then
    echo "Creating .env file from template..."
    cp "backend/.env.template" "backend/.env"
    echo
    echo "Please edit backend/.env file and add your DASHSCOPE_API_KEY"
    echo "Then run this script again."
    exit 1
fi

# Start the application
echo "Starting backend server..."
cd backend
python run.py
