How-To
 1. install [python (3.11.x)](https://www.python.org/downloads/) 
 2. initial the python enviroment with command (run in powershell or cmd terminal): 
 ```
 pip install -r requirements.txt
 ```
3. put the "Extractor.py" and "template.binary" to the root folder which contains all the projects
4. open a powershell (or cmd), and cd to the root folder of the projects
5. run the command to extract the project subtotal information:
```
python Extractor.py
```

it will create a new file "ExtractResult.xlsx", which contains all the subtotal info you need.