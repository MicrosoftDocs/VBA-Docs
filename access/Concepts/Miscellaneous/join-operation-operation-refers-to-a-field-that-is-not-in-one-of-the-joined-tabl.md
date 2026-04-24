---
title: JOIN operation <operation> refers to a field that is not in one of the joined tables. (Error 3082)
keywords: jeterr40.chm5003082
f1_keywords:
- jeterr40.chm5003082
ms.assetid: 13a1b996-709e-198a-fe68-9a23fd39f6a7
ms.date: 06/08/2019
ms.localizationpriority: medium
---


# JOIN operation \<operation\> refers to a field that is not in one of the joined tables. (Error 3082)

  

**Applies to:** Access 2013 | Access 2016

You tried to execute an SQL statement that contains an invalid join. This error occurs when you try to create a join on a field that is not in one of the joined tables. The following example produces this error.




```sql
SELECT Authors.FirstName, Titles.ISBN 
FROM Authors, Titles, AuthorTitles, 
Authors INNER JOIN Titles ON Authors.ID = AuthorTitles.ISBN;
```

The error occurs because the join involves the Authors and Titles tables, but the joined fields are in the Authors and AuthorTitles tables.

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]