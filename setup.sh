#!/bin/bash

cd ./setup/
rm -r openpyxl-2.6.3.tar.gz openpyxl-2.6.3
powershell.exe -File ./get_openpyxl.ps1
./unzip_openpyxl.sh
powershell.exe -File ./install_openpyxl.ps1
rm -r openpyxl-2.6.3.tar.gz openpyxl-2.6.3
