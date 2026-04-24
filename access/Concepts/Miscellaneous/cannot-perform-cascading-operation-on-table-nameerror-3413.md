---
title: Cannot perform cascading operation on table <name>. (Error 3413)
keywords: jeterr40.chm5003413
f1_keywords:
- jeterr40.chm5003413
ms.assetid: 0a060fbf-86a5-24b4-8a46-2f2ea90ea4ab
ms.date: 06/08/2019
ms.localizationpriority: medium
---


# Cannot perform cascading operation on table \<name\>. (Error 3413)

  

**Applies to:** Access 2013 | Access 2016

This error occurs when referential integrity is defined with one of the following actions: CASCADE DELETE, UPDATE or NULL. The error occurs when making a change to a table that contains a referenced primary key and one or more of the foreign key tables is opened in a mode that prevents the foreign key tables from being opened in a shared read/write mode.

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](https://learn.microsoft.com/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]