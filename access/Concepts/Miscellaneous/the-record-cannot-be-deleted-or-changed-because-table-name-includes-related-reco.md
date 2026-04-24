---
title: The record cannot be deleted or changed because table <name> includes related records. (Error 3200)
keywords: jeterr40.chm5003200
f1_keywords:
- jeterr40.chm5003200
ms.assetid: e3171406-6a42-5932-35f4-b0a4db616f3a
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# The record cannot be deleted or changed because table \<name\> includes related records. (Error 3200)

  

**Applies to:** Access 2013 | Access 2016

You tried to perform an operation that would have violated referential integrity rules for related tables. For example, this error occurs if you try to delete or change a record in the "one" table in a one-to-many relationship when there are related records in the "many" table.

If you want to delete or change the record, first delete the related records from the "many" table.

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](https://learn.microsoft.com/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
