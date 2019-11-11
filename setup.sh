#!/bin/bash

cd ./setup/
./remove_openpyxl.sh
powershell.exe -File ./get_openpyxl.ps1
./unzip_openpyxl.sh
powershell.exe -File ./install_openpyxl.ps1
./remove_openpyxl.sh
