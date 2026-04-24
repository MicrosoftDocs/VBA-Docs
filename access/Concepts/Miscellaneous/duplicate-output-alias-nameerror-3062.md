---
title: Duplicate output alias <name>. (Error 3062)
keywords: jeterr40.chm5003062
f1_keywords:
- jeterr40.chm5003062
ms.assetid: e0157e7c-d854-4a9a-b5ba-22afa0944cbc
ms.date: 06/08/2019
ms.localizationpriority: medium
---


# Duplicate output alias \<name\>. (Error 3062)

  

**Applies to:** Access 2013 | Access 2016

You tried to execute an SQL statement that has more than one alias with the same name. The following statement, for example, would produce this error:




```sql
SELECT LastName AS Name, FirstName AS Name FROM Employees;

```

Rename one or more of the aliases, and then try the operation again.

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](https://learn.microsoft.com/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
