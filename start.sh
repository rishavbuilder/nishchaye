#!/bin/bash
cd /home/rishav/Downloads/nishchay

# Create virtual env if not exists
if [ ! -d "/tmp/nikshay_env" ]; then
    echo "Creating virtual environment..."
    python3 -m venv /tmp/nikshay_env
fi

# Install/upgrade dependencies
echo "Installing dependencies..."
/tmp/nikshay_env/bin/pip install --quiet --upgrade selenium webdriver-manager openpyxl python-dotenv

# Run the bot
echo "Starting Nikshay Bot..."
/tmp/nikshay_env/bin/python3 nikshay_data_entry.py

read -p "Press Enter to close..."
