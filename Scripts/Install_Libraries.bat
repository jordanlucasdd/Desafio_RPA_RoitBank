@echo off

:start
cls

set python_ver=38

cd \
cd \python%python_ver%\Scripts\
pip install openpyxL
pip install pandas
pip install xlsxwriter
pip install unidecode


pause
exit