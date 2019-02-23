---
title: ODBCConnectStr property
ROBOTS: INDEX
keywords: vbaac10.chm4442
f1_keywords:
- vbaac10.chm4442
ms.prod: access
api_name:
- Access.ODBCConnectStr
ms.assetid: f1eba90d-ec30-7e71-a0ca-0d8ed81ac61b
ms.date: 06/08/2017
localization_priority: Normal
---


# ODBCConnectStr property

**Applies to:** Access 2013 | Access 2016

You can use the **ODBCConnectStr** property in an SQL pass-through query to specify the Open Database Connectivity (ODBC) connection string for the query.

> [!NOTE] 
> The **ODBCConnectStr** property applies only to pass-through queries.

## Setting

Enter the ODBC connection string that defines the connection to the SQL database you want to use.

You can set this property by using the query's property sheet or Visual Basic.

You can also use the ODBC Connection String Builder to create the ODBC connection string for this property. This builder establishes a connection to the SQL database server and then ends the connection after the ODBC connection string is created.

## Remarks

The default setting for this property is the string "ODBC;", which Microsoft Access restores if you delete an existing setting. When the **ODBCConnectStr** property is set to "ODBC;", Microsoft Access will prompt you for a connection string whenever the query is used but won't store the connect string. You must enter a connection string in the **ODBCConnectStr** property box if you want Microsoft Access to store the connection string.

> [!TIP] 
> If you know the full connection string for the SQL database, enter it in the **ODBCConnectStr** property box. This way you will avoid the need to enter the connection string in the ODBC connection dialog box each time you use the query.

The connection string is different for different types of ODBC data sources. For example, to connect to the Human Resources data source on the HRSRVR server (a Microsoft SQL Server) by using the logon identification (ID) Smith and the password Sesame, you can use the following connection string.

    ODBC;DSN=Human Resources;SERVER=HRSRVR;UID=Smith;PWD=Sesame;

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]