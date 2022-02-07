# Description
Tax season come again. It's headache when calculating PL(profit or loss) of your investment or business manually, especially for people with big volume of trading records. Now I developed a Python program to calculate PL automatically.

# File description
1 trade_sample.txt: The file you get from your trading company, it's the trading records. Some trading company's file may be different from this sample that require changing Python codes to fit.<br/>
2 Account_Sample.xls: The file you submit to your tax bureau. It included the PL data which is used to calculate your income.<br/>
3 taxCalculate_Sample.py: this is the Python code to calculate PL. The python code first initial the last year's position data, then read trading records in trade_sample.txt, then wirte the file of Account_Sample.xls to tell PL.<br/>

# Run Guide
1 Install python-3.7.9, and MS-Office<br/>
2 Run command of "pip install xlwings" to install python library for Excel.<br/>
3 Use CD command to change your current path to this project.<br/>
4 Run command of "python taxCalculate_Sample.py" to calculate PL.<br/>
5 You also can use MS Visual Studio Code to run or debug this python codes.
