#!/bin/bash

echo "Adding path.txt files"
cd src/
SRC=$(pwd)
echo "$SRC" > path.txt
echo "$SRC" > ../path.txt
echo "$SRC" > ../setup/path.txt
echo "Added path.txt files"
cd ../
echo "Installing latest version of pip"
python3 -m pip install --user --upgrade pip
echo "Installing virtualenv"
python3 -m pip install --user virtualenv
echo "Setting up virtual environment"
python3 -m venv env