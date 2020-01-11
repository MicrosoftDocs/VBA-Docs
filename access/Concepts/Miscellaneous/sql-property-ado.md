---
title: SQL property [ADO]
ROBOTS: INDEX
ms.prod: access
ms.assetid: 210adcbb-5c89-150b-4c61-6a52dea9af56
ms.date: 06/08/2017
localization_priority: Normal
---


# SQL property [ADO]

**Applies to:** Access 2013 | Access 2016

Indicates the query string used to retrieve the [Recordset](https://msdn.microsoft.com/library/0f963bf8-f066-dc8a-b754-f427de712df1%28Office.15%29.aspx).

You can set the **SQL** property at design time in the [RDS.DataControl](https://msdn.microsoft.com/library/ac430669-7628-696c-c036-b5d35405d788%28Office.15%29.aspx) object's OBJECT tags, or at run time in scripting code.

## Parameters

-  _QueryString_
    
    - A **String** value that contains a valid SQL data request.
    
-  _DataControl_
    
    - An object variable that represents an **RDS.DataControl** object.
    

## Remarks

In general, this is an SQL statement (using the dialect of the database server), such as `.` To ensure that records are matched and updated accurately, an updatable query must contain a field other than a Long Binary field or a computed field.

The **SQL** property is optional if a custom server-side business object retrieves the data for the client.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]