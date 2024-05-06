#!/bin/bash

echo -e "Checking if the script is running as root!"

# Check if script is running as root user
if [[ $EUID -ne 0 ]]; then
    echo "This script must be run as root. Please use 'sudo' along with the command or login as root user."
    sleep 4
    echo "Saving Preferences for next try!"
    sleep 7
    exit 1
fi

# Check if Colorama has been already installed or not
if python3 -c "import colorama" &>/dev/null; then
    echo -e "\033[95mColorama has been already installed, We have initialized it for you :)\033[0m"
    sleep 5
else
    echo -e "\033[91mColorama has not been installed. Installing it...\033[0m"
    pip install colorama
    echo -e "\033[95mDone, Colorama has been installed.\033[0m"
    sleep 3
fi


echo -e "\033[93m\nUpdating your system, wait!\n\033[0m"
sleep 0.5

# Update system
apt update

# Install pip
apt install python3-pip -y

echo -e "\033[96m\nInstalling required dependencies, wait!\n\033[0m"
sleep 0.5

# Install requirements to run the script
sudo apt-get install libreoffice -y

echo -e "\033[32m\nDone, starting next task now!\n\033[0m\n"
sleep 2

# Prompt the user to input the path of the CSV file
read -p "\033[95m\nEnter the path of the CSV file: \033[0m" csv_file

# Check if the CSV file exists
if [ ! -f "$csv_file" ]; then
    echo "\033[91m\nError: CSV file not found.\033[0m\n"
    exit 1
fi

# Path to the Excel file
excel_file="output.xlsx"

# Check if LibreOffice is already running
if pgrep -x "soffice" > /dev/null
then
    echo "\033[92m\nLibreOffice is already running.\033[0m\n"
else
    echo "\033[91m\nLibreOffice is not running. Opening LibreOffice...\033[0m\n"
    libreoffice --calc &
    sleep 9 # Wait for LibreOffice to open
fi

# Convert CSV to Excel
libreoffice --convert-to xlsx "$csv_file" --outdir "$(pwd)"

# Wait for LibreOffice to finish converting
sleep 9

# Open the Excel file
libreoffice "$excel_file"
