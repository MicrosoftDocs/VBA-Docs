---
title: Circular reference caused by <query reference>. (Error 3102)
keywords: jeterr40.chm5003102
f1_keywords:
- jeterr40.chm5003102
ms.prod: access
ms.assetid: f3ff3f1b-8a2f-7038-57f8-0abadde1c3cf
ms.date: 06/08/2019
localization_priority: Normal
---


# Circular reference caused by <query reference>. (Error 3102)

  

**Applies to:** Access 2013 | Access 2016

You tried to execute a query that depends on itself for data. For example, this error occurs if you execute either of the following queries:

Query1



```sql
SELECT * FROM Employees, Query2;

```

Query2



```sql
SELECT * FROM Query1;


```

Redesign the queries to eliminate the dependency.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
