---
title: Update/delete conflict - This updated record was deleted at another replica. (Error 3736)
keywords: jeterr40.chm5003736
f1_keywords:
- jeterr40.chm5003736
ms.assetid: d8e66115-9a71-72b1-137b-61305057fb00
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# Update/delete conflict - This updated record was deleted at another replica. (Error 3736)

  

**Applies to:** Access 2013 | Access 2016

When a record is deleted at one replica, but updated at another replica, the deleted record always wins in the conflict that occurs when the two replicas synchronize. The updated record is logged in the conflict table. To reverse the initial resolution of the conflict, reinsert the conflict record. To accept the current resolution, delete the conflict record.

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]