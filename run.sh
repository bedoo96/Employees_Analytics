#!/bin/bash

# Employee Attendance Analytics - Startup Script

echo "🚀 Starting Employee Attendance Analytics System..."
echo ""

# Check if virtual environment exists
if [ ! -d "venv" ]; then
    echo "📦 Creating virtual environment..."
    python3 -m venv venv
fi

# Activate virtual environment
echo "🔧 Activating virtual environment..."
source venv/bin/activate

# Install/upgrade dependencies
echo "📥 Installing dependencies..."
pip install -r requirements.txt --quiet

# Check if .env exists
if [ ! -f ".env" ]; then
    echo "⚠️  Warning: .env file not found!"
    echo "   Please copy .env.example to .env and configure your API keys"
    echo ""
fi

# Clear the terminal
clear

# Start the application
echo "✨ Launching Streamlit application..."
echo ""
echo "📊 Employee Attendance Analytics System"
echo "   Access the application at: http://localhost:8501"
echo ""
echo "   Press Ctrl+C to stop the server"
echo ""

streamlit run app0.py