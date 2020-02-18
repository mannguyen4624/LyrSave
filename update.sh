echo '.vs/' > .gitignore
echo 'path.txt' >> .gitignore

git add lyr.xlsx
git commit -m "Update spreadsheet"
git push

echo 'lyr.xlsx' >> .gitignore
