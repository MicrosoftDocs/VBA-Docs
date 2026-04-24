---
title: Circular reference caused by alias <name> in query definition's SELECT list. (Error 3103)
keywords: jeterr40.chm5003103
f1_keywords:
- jeterr40.chm5003103
ms.assetid: b0f5d8a6-4735-367f-dd27-af3d97816430
ms.date: 06/08/2019
ms.localizationpriority: medium
---


# Circular reference caused by alias \<name\> in query definition's SELECT list. (Error 3103)

  

**Applies to:** Access 2013 | Access 2016

The specified alias created a reference that cannot be resolved. This error can occur, for example, if you enter the following SQL statement, in which A is the circular reference:




```sql
SELECT A + B AS C, C + D AS E, E + F AS A

FROM MyTable;
```




```sql
SELECT week1 + week2 as hours, hours + overtime as gross, gross + ytdpay as week1 FROM EmployeePay

```

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](https://learn.microsoft.com/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
