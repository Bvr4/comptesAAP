@echo off


echo   --   lancement de libre office en mode ecoute   --
echo   /!\  Cela peut peut prendre quelques secondes
echo        Si cela dure plusieurs minutes, merci de relancer le programme

"C:\\Program Files\LibreOffice 5\program\soffice.exe" --nodefault --accept="socket,host=localhost,port=2002;urp;

"
cd scripts

echo   --   lancement du script de creation de nouveau fichier   --

"C:\\Program Files\LibreOffice 5\program\python.exe" creerFichier.1.0.py

echo   ------------------------------------------------------
echo   --   Vous pouvez mainetenant fermer cette fenetre   --
echo   ------------------------------------------------------
pause >nul