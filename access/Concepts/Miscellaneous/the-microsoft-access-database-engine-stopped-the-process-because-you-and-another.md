---
title: The Microsoft Access database engine stopped the process because you and another user are attempting to change the same data at the same time. (Error 3197)
keywords: jeterr40.chm5003197
f1_keywords:
- jeterr40.chm5003197
ms.assetid: 3ea30548-166c-2cfc-5014-6d624a75294e
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# The Microsoft Access database engine stopped the process because you and another user are attempting to change the same data at the same time. (Error 3197)

  

**Applies to:** Access 2013 | Access 2016

This error can occur in a multiuser environment.

Another user has changed the data you are trying to update. This error can occur when multiple users open a table or create a **Recordset** and use optimistic locking. Between the time you used the **Edit** method and the **Update** method, another user changed the same data.
To overwrite the other user's changes with your own, execute the **Update** method again.

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
