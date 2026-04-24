---
title: ODBCTimeout property
ROBOTS: INDEX
keywords: vbaac10.chm4443
f1_keywords:
- vbaac10.chm4443
api_name:
- Access.ODBCTimeout
ms.assetid: ebcac9df-87a9-481c-32cc-d28bb9f37717
ms.date: 06/08/2019
ms.localizationpriority: medium
---


# ODBCTimeout property

**Applies to:** Access 2013 | Access 2016

Use the **ODBCTimeout** property to specify the number of seconds Microsoft Access waits before a time-out error occurs when a query is run on an Open Database Connectivity (ODBC) database.


## Setting

The **ODBCTimeout** property is an Integer value representing the number of seconds Microsoft Access waits. The default is 60 seconds. When this property is set to 0, no time-out error occurs.

You can set this property by using the query's property sheet or Data Access Objects (DAO) in Visual Basic code.


## Remarks

When you are using an ODBC database, such as Microsoft SQL Server, there may be delays due to network traffic or heavy use of the ODBC server. The **ODBCTimeout** property lets you specify how long Microsoft Access waits for a network connection before a time-out error occurs.

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](https://learn.microsoft.com/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]