#!/bin/bash

touch path.txt
echo $PWD > path.txt
cd ./setup/
./remove_openpyxl.sh
powershell.exe -File ./get_openpyxl.ps1
./unzip_openpyxl.sh
./install_openpyxl.sh
./remove_openpyxl.sh

./resetexcel.py
