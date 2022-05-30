dir /b lotes > saidas.txt
(for /f "Tokens=1" %%a in (saidas.txt) do par.bat %%a  )
