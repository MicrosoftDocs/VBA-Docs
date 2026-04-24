---
title: Table <name> already exists. (Error 3010)
keywords: jeterr40.chm5003010
f1_keywords:
- jeterr40.chm5003010
ms.assetid: eaf04a4d-15cb-b27a-d6e3-8f89e88d4143
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# Table \<name\> already exists. (Error 3010)

  

**Applies to:** Access 2013 | Access 2016

You tried to create or rename a table with a name that already exists in this database. Choose another name, and then try the operation again.

In a multiuser database, this error can also occur if you delete a table, another user creates a table with the same name, and then you try to roll back the deletion of your table. To restore your table, the other user must first delete or rename the new table before you try the rollback operation again.

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](https://learn.microsoft.com/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]