---
title: Specify the table containing the records you want to delete. (Error 3128)
keywords: jeterr40.chm5003128
f1_keywords:
- jeterr40.chm5003128
ms.assetid: f6c49cba-5b9c-775c-625a-6d1e79c8adf0
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# Specify the table containing the records you want to delete. (Error 3128)

**Applies to:** Access 2013 | Access 2016

You tried to execute a delete query but the query does not specify the name of the table containing the records you want to delete.

Possible cause:

- You did not type an asterisk for each table in the ALL, DISTINCT, DISTINCTROW predicates. Instead, you typed field names (for example, `Customers.Address` instead of `Customers.*`).

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](https://learn.microsoft.com/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
