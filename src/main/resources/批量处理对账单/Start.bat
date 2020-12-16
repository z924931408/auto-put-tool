@echo off 
copy /y CopyName.bat 对账单
copy /y 对账单批量名字集合.txt 对账单
cd D:\批量处理对账单\对账单
call CopyName.bat

java -jar D:\批量处理对账单\Auto.jar

