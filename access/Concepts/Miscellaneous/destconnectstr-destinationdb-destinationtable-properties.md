---
title: DestConnectStr, DestinationDB, DestinationTable properties
ROBOTS: INDEX
keywords: vbaac10.chm6181
f1_keywords:
- vbaac10.chm6181
ms.prod: access
ms.assetid: 5d9f3e9d-fc36-d4af-d82b-8d2ebd8044fa
ms.date: 06/08/2017
localization_priority: Normal
---


# DestConnectStr, DestinationDB, DestinationTable properties

**Applies to:** Access 2013 | Access 2016

- The **DestinationDB** property specifies the type of application used to create an external database.
    
- The **DestConnectStr** property specifies the connection string for the name of the database that will contain the new table (for make-table queries) or the name of the database that contains the table to which data will be appended (for append queries).
    
- The **DestinationTable** property specifies the name of the table that will hold the results of the make-table or append query.
    

> [!NOTE] 
> The **DestConnectStr**, **DestinationDB**, and **DestinationTable** properties apply only to make-table and append queries.


## Setting

You set the **DestConnectStr**, **DestinationDB**, and **DestinationTable** properties by using a string expression.

The default setting for **DestinationDB** property is "(current)", which refers to the currently active database.

You can set these properties in the query's property sheet, or in SQL view of the Query window.

In the SQL statement for an append query, the table name in the INSERT INTO statement corresponds to the **DestinationTable** property setting. The IN clause corresponds to the **DestinationDB** and **DestConnectStr** property settings.

In the SQL statement for a make-table query, the table name in the INTO clause corresponds to the **DestinationTable** property setting. The IN clause corresponds to the **DestinationDB** and **DestConnectStr** property settings.

> [!NOTE] 
> Microsoft Access sets these properties automatically based on the information you enter in the query's property sheet or in SQL view of the Query window.


## Remarks

When you click **Make Table** or **Append** on the **Query** menu, Microsoft Access prompts you for the information needed to set these properties. Microsoft Access uses the value you enter in the **Table Name** box to set the **DestinationTable** property, and it uses the information you type in the **File Name** box to set the **DestConnectStr** and **DestinationDB** properties.

To use the query's property sheet to specify a table in a Microsoft Access database, enter the name of the table in the **DestinationTable** property box (for example, Clients). In the **DestinationDB** property box, enter the path and database name without the file name extension (for example, C:\Accounts\Customers). Microsoft Access adds the extension automatically. You don't need to set the **DestConnectStr** property.

To specify a table in a database created by a different product, such as Paradox, enter the name of the table in the **DestinationTable** property box. In the **DestinationDB** property box, enter the path (for example, C:\Pdoxdata). In the **DestConnectStr** property box, enter the specifier for the database type (for example, Paradox 3.x). For a list of specifiers, see the ADO **Connect** property.

To specify an Open Database Connectivity (ODBC) database, enter the name of the database in the **DestConnectStr** property box along with additional information, such as a logon identification (ID) and password, as required by the product. For example, you might use the following setting for a Microsoft SQL Server database:

    ODBC;DSN=salessrv;UID=jace;PWD=password;DATABASE=sales;

For more information about ODBC drivers, such as Microsoft SQL Server, see the Help provided with the driver.

You don't need to set the **DestinationDB** property for ODBC databases.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]