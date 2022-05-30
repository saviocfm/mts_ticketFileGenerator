rem echo %1
move ./lotes/%1/xml ./
move ./lotes/%1/xlsx ./ 
python app.py 
move  ./xml ./lotes/%1
move  ./xlsx ./lotes/%1