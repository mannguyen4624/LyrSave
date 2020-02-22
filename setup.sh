#!/bin/bash

cd src/
SRC=$(pwd)
echo "$SRC" > path.txt
echo "$SRC" > ../path.txt
echo "$SRC" > ../setup/path.txt
cd ../setup/
python3 -m venv env

