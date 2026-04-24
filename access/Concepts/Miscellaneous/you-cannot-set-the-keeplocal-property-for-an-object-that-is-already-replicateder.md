---
title: You cannot set the KeepLocal property for an object that is already replicated. (Error 3457)
keywords: jeterr40.chm5003457
f1_keywords:
- jeterr40.chm5003457
ms.assetid: 916ea4af-3190-99f4-901d-76b7754efa6a
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# You cannot set the KeepLocal property for an object that is already replicated. (Error 3457)

  

**Applies to:** Access 2013 | Access 2016

The **KeepLocal** property cannot be set on a replicated object. Setting a local object's **KeepLocal** property after the database has been replicated has no effect on the object. If you want to keep an object from being replicated to the other replicas in the set, set the object's **Replicable** property to "F".

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](https://learn.microsoft.com/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]