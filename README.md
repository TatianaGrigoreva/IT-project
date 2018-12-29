# IT-project
ПО необходимое для запуска скриптов, импорта БД: Python версии 3+, MS SQL Server, Jupyter Notebook.

Ссылка на скачивание Python: https://www.python.org/downloads/

Установка Jupyter Notebook осуществляется из cmd:

python3 -m pip install --upgrade pip

python3 -m pip install jupyter

Запуск Jupyter: jupyter notebook

Необходимые библиотеки (устанавливаются также запросами из cmd):

-pip install WeasyPrint; -pip install pyodbc; -pip install psycopg2; -pip install cairocffi; -pip install sqlalchemy; -pip install jinja2; -pip install matplotlib; -pip install pandas


До запуска кода важно поместить данные (файлы base_prices.xlsx, bond_description.xlsx, cbr_rates.xlsx) в одну папку с файлами c расширением .ipynb. Дальнейшие инструкции содержатся непосредственно в скриптах.
