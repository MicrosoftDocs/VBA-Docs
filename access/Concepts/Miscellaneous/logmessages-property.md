---
title: LogMessages property
ROBOTS: INDEX
keywords: vbaac10.chm5187437
f1_keywords:
- vbaac10.chm5187437
ms.prod: access
api_name:
- Access.LogMessages
ms.assetid: 848f215b-50aa-22f4-264c-ff7d00347aa7
ms.date: 06/08/2017
localization_priority: Normal
---


# LogMessages property

**Applies to:** Access 2013 | Access 2016

You can use the **LogMessages** property in an SQL pass-through query to specify whether messages returned from an SQL database are stored in a messages table in the current Microsoft Access database.

> [!NOTE] 
> The **LogMessages** property applies only to pass-through queries.


## Setting

The **LogMessages** property uses the following settings.

|Setting|Visual Basic|Description|
|:-----|:-----|:-----|
|Yes|**True** (-1)|Microsoft Access stores messages returned from the SQL database in a messages table.|
|No|**False** (0)|(Default) Microsoft Access doesn't store messages returned from the SQL database.|

You can set this property by using the query's property sheet or Visual Basic.


## Remarks

The name of the messages table where the returned messages are stored is  _username - nn_, where _username_ is the logon name of the user running the pass-through query, and _nn_ is an integer that increases in increments of 1, starting at 00. For example, if user JoanW sets the **LogMessages** property to Yes and receives messages from an SQL database, the messages table will be named JoanW - 00. If JoanW receives messages in another Microsoft Access session (and the first table hasn't been deleted), a new table named JoanW - 01 is created.

> [!NOTE] 
> Error messages from SQL Server aren't stored in the messages table.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]