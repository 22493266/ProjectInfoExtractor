# How-To
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


# How-To
## 环境安装
1. 安装[python (3.11.x)](https://www.python.org/downloads/) ，在下载页面下载3.11.*的版本
2. 使用下面的这个命令去初始化本地python运行环境（打开powershell或者cmd，切换到本项目解压的目录下）
 ```
 pip install -r requirements.txt
 ```
 ## 如何使用
 1. 将Extractor.py 和 template.binary 放到想要提取的项目集合的根目录
 2. 在该根目录打开一个powershell或者cmd命令行终端，并执行以下命令：
 ```
python Extractor.py
```

执行完毕之后，在同一目录下会生成一个"ExtractResult.xlsx"，这个就是汇总信息。
