---
title: Parameter <name> specified where a database name is required. (Error 3217)
keywords: jeterr40.chm5003217
f1_keywords:
- jeterr40.chm5003217
ms.assetid: d6d700f2-5df5-5d26-a6ee-706ca4c1a12a
ms.date: 06/08/2019
ms.localizationpriority: medium
---


# Parameter \<name\> specified where a database name is required. (Error 3217)

  

**Applies to:** Access 2013 | Access 2016

You created a parameter query that specifies an invalid parameter type. The following example produces this error:




```sql
PARAMETERS Param1 Text;

SELECT CustomerID
FROM Customers IN Param1;
```

 `Param1` is a text parameter, but the FROM clause requires a database parameter.
Change the parameter type from Text to Database, and then try the operation again.

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]