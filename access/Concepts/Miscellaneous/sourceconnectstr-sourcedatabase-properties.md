---
title: SourceConnectStr, SourceDatabase properties
ROBOTS: INDEX
keywords: vbaac10.chm4548
f1_keywords:
- vbaac10.chm4548
ms.prod: access
ms.assetid: eed57130-f030-b800-5b1a-92249d6c23a5
ms.date: 06/08/2017
localization_priority: Normal
---


# SourceConnectStr, SourceDatabase properties

**Applies to:** Access 2013 | Access 2016

You can use these properties to access external data when you can't link the external tables to your database.

- The **SourceConnectStr** property specifies the name of the application used to create an external database.
    
- The **SourceDatabase** property specifies the external database in which the source tables or queries for a query reside.
    
> [!NOTE] 
> The **SourceConnectStr** and **SourceDatabase** properties apply to all queries except data-definition, pass-through, and union queries.


## Setting

You use a string expression to set the value of the **SourceConnectStr** and **SourceDatabase** properties.

You can set these properties in the query's property sheet or in SQL view of the Query window. In the SQL statement, the properties correspond to the IN clause.

> [!NOTE] 
> If you are accessing multiple database sources, use the **Source** property instead of the **SourceConnectStr** and **SourceDatabase** properties.


## Remarks

You must use the **SourceConnectStr** and **SourceDatabase** properties to access tables from external databases that were created in applications that don't use linked tables.

The following are examples of these property settings:

- For a Microsoft Access database, the **SourceDatabase** property setting is the path and database name (for example, C:\Accounts\Customers). Microsoft Access adds the file name extension automatically. The **SourceConnectStr** property doesn't have a value for a Microsoft Access database.
    
- For a database created by a product such as Paradox, the **SourceDatabase** property setting is the path (for example, C:\Pdoxdata). The **SourceConnectStr** property setting is the specifier for the database type (for example, Paradox 3.x;). For a list of specifiers, see the DAO **Connect** property.
    
- The following example uses dBASE IV tables in the C:\Dbdata directory as the source of data for the query.
    
  ```sql
    SELECT Customer.COMPANYNAM, Orders.ORDERID, Orders.ORDERDATE 
  FROM Customer INNER JOIN Orders 
  ON Customer.CUSTOMERID = Orders.CUSTOMERID 
  IN 'C:\Dbdata'[dBASE IV;];
  ```

- For an Open Database Connectivity (ODBC) database, the **SourceConnectStr** property setting is the name of the source database and any additional information required by the product, such as a logon identification (ID) and password. For example, for a Microsoft SQL Server database the setting might be:
    

    ODBC;DSN=salessrv;UID=jace;PWD=password;DATABASE=sales;



The **SourceDatabase** property doesn't have a value for an ODBC database.

## See also

- [DataControl Object (RDS)](https://msdn.microsoft.com/library/ac430669-7628-696c-c036-b5d35405d788%28Office.15%29.aspx) - [Server Property (RDS)](https://msdn.microsoft.com/library/17519dbe-a43a-1d0d-22c1-dc0def2f63ab%28Office.15%29.aspx)
- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]