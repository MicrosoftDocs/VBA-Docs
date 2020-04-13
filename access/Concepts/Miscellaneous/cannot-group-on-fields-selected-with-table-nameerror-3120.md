---
title: Cannot group on fields selected with '*' <table name>. (Error 3120)
keywords: jeterr40.chm5003120
f1_keywords:
- jeterr40.chm5003120
ms.prod: access
ms.assetid: 34cce8ec-dc95-7f1d-8537-9dd7dbbc442d
ms.date: 06/08/2019
localization_priority: Normal
---


# Cannot group on fields selected with '*' <table name>. (Error 3120)

  

**Applies to:** Access 2013 | Access 2016

You tried to execute a SELECT statement that groups or totals all fields in a single table, selected with an asterisk ( * ). This error occurs, for example, if you enter the following SQL statement:




```sql
SELECT Orders.* FROM Orders GROUP BY ShipVia;

```

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]