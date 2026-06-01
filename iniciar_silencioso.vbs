Dim WshShell
Set WshShell = CreateObject("WScript.Shell")

' Aguarda 10 segundos para o OneDrive e a rede carregarem
WScript.Sleep 10000

Dim pasta
pasta = "C:\Users\vitho\OneDrive\READET~1\ABRIL2~4\SISTEM~1"

Dim python
python = "C:\Users\vitho\AppData\Local\Python\PYTHON~1.14-\python.exe"

WshShell.CurrentDirectory = pasta
WshShell.Run """" & python & """ app.py", 0, False
