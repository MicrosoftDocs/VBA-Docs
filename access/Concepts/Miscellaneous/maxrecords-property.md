---
title: MaxRecords property
ROBOTS: INDEX
keywords: vbaac10.chm4422
f1_keywords:
- vbaac10.chm4422
api_name:
- Access.MaxRecords
ms.assetid: 30ea62b8-9304-2cdf-ff2f-d8ed665b16b4
ms.date: 06/08/2019
ms.localizationpriority: medium
---


# MaxRecords property

**Applies to:** Access 2013 | Access 2016

Specifies the maximum number of records that will be returned by:

- A query that returns data from an ODBC database to a Microsoft Access database. 
    
- A view that returns data from a SQL database to an Access project (.adp). 
    

## Setting

The **MaxRecords** property setting is a Long Integer value representing the number of records that will be returned.

In a Microsoft Access database, you can set this property by using the query's property sheet or Visual Basic.


## Remarks

When you set this property in Visual Basic, you use the ADO **MaxRecords** property.

Records are returned in the order specified by the query's ORDER BY clause.

Use the **MaxRecords** property in situations where limited system resources might prohibit a large number of returned records.

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](https://learn.microsoft.com/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]