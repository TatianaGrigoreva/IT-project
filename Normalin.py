# -*- coding: utf-8 -*-
"""
Created on Sun Dec 30 14:51:02 2018

@author: mikha
"""

import pyodbc

server = 'DESKTOP-HP138CN\SQLEXPRESS' #server name
database = 'Risks' 

connection = pyodbc.connect('Driver={SQL Server};SERVER=' + server 
                            + ';DATABASE='+ database
                            + ';Trusted_Connection=yes') #create new connection with server
cursor = connection.cursor() #create special object for queries			

cursor.execute("""SELECT * INTO [dbo].[Base_prices] FROM [dbo].[Base_1] UNION ALL SELECT * FROM [dbo].[Base_2]""") # union two tables (Base_1 and Base_2)
connection.commit()

cursor.execute("""ALTER TABLE [dbo].[Base_prices] ALTER COLUMN [ID] varchar(80) NOT NULL""")
connection.commit()

cursor.execute("""ALTER TABLE [dbo].[Base_prices] ALTER COLUMN [ISIN] varchar(80) NOT NULL""")
connection.commit()

cursor.execute("""ALTER TABLE [dbo].[Bond] ALTER COLUMN [ISIN_RegCode_NRDCode] nvarchar(200) NOT NULL""")
connection.commit()

cursor.execute("""ALTER TABLE [dbo].[Instrs] ALTER COLUMN [ID] nvarchar(200) NOT NULL""")
connection.commit()

cursor.execute("""ALTER TABLE [dbo].[Instrs] ADD PRIMARY KEY (ID)""")
connection.commit()

cursor.execute("""ALTER TABLE [dbo].[Bond] ADD PRIMARY KEY (ISIN_RegCode_NRDCode)""")
connection.commit()

cursor.execute("""ALTER TABLE [dbo].[Base_prices] ADD CONSTRAINT FK_Base_1 FOREIGN KEY (ID) REFERENCES [dbo].[instrs] (ID)""")
connection.commit()

cursor.execute("""SELECT [dbo].[Base_prices].[ID] INTO [dbo].[Uknown_ISIN_ID] FROM [dbo].[Base_prices] LEFT JOIN [dbo].[Bond] ON [Base_prices].[ISIN] = [Bond].[ISIN_RegCode_NRDCode] WHERE [Bond].[ISIN_RegCode_NRDCode] IS NULL""")
connection.commit()

cursor.execute("""SELECT * INTO [dbo].[Uknown_ISIN_Prices] FROM [dbo].[Base_prices] WHERE (ID) IN (SELECT (ID) FROM [dbo].[Uknown_ISIN_ID]""")
connection.commit()

cursor.execute("""DELETE FROM [dbo].[base_prices] WHERE [ID] IN (SELECT (ID) FROM [dbo].[Unknown_ISIN_ID]""")
connection.commit()

cursor.execute("""ALTER TABLE [Base_prices] ADD CONSTRAINT FK_Base_2 FOREIGN KEY (ISIN) REFERENCES [dbo].[Bond] [ISIN_RegCode_NRDCode]""")
connection.commit()

cursor.execute("""SELECT COUNT(*)*100/(SELECT COUNT(*) FROM [dbo].[Bond]) FROM [dbo].[Bond] WHERE (ISIN144A)=' ' """)
connection.commit()

cursor.execute("""SELECT (ISIN_RegCode_NRDCode), (ISIN144A) INTO [dbo].[Bond_ISIN144A] FROM [dbo].[Bond] WHERE (ISIN144A) !=' ' """)
connection.commit()

cursor.execute("""ALTER TABLE [dbo].[Bond] DROP COLUMN (ISIN144A)""")
connection.commit()

cursor.close()
connection.close()