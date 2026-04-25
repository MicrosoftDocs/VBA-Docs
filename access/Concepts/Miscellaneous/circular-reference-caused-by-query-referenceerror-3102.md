---
title: Circular reference caused by <query reference>. (Error 3102)
keywords: jeterr40.chm5003102
f1_keywords:
- jeterr40.chm5003102
ms.assetid: f3ff3f1b-8a2f-7038-57f8-0abadde1c3cf
ms.date: 06/08/2019
ms.localizationpriority: medium
---


# Circular reference caused by \<query reference\>. (Error 3102)

  

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

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
