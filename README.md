Financial Data Scraper
===================


#### Project Preparation
------------------------------------------------------------------------

Goal
-------------------
 - Acquire historical financial data for ASX listed companies
 - Automate the process

----------

Process
-------------
Located a data provider that provided the following button to download all financial information (historical included) to excel.

![Capture](https://user-images.githubusercontent.com/43980002/67251491-0fa45380-f4bb-11e9-89e9-2bcb04befbc5.JPG)

First part of the program iterates through all ASX codes, updating the URL and saving all .xls files. The next part of the code converts all .xls files into .xlsx and removes the original .xls file.

![Cap](https://user-images.githubusercontent.com/43980002/67251759-dcae8f80-f4bb-11e9-8eca-6daf822e086e.JPG)

Finally, we want to go from the example about for each company and combine all the same statements into their own excel file i.e. P&L.xlsx, balance_sheet.xlsx etc. 

----------

Requirements
--------------------
- urllib  
- pandas  
- win32com   
- openpyxl  
- os  
----------
