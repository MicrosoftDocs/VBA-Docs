---
title: Cannot group on fields selected with '*' <table name>. (Error 3120)
keywords: jeterr40.chm5003120
f1_keywords:
- jeterr40.chm5003120
ms.assetid: 34cce8ec-dc95-7f1d-8537-9dd7dbbc442d
ms.date: 06/08/2019
ms.localizationpriority: medium
---


# Cannot group on fields selected with '*' <table name>. (Error 3120)

  

**Applies to:** Access 2013 | Access 2016

You tried to execute a SELECT statement that groups or totals all fields in a single table, selected with an asterisk ( * ). This error occurs, for example, if you enter the following SQL statement:




```sql
SELECT Orders.* FROM Orders GROUP BY ShipVia;

```

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]