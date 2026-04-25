---
title: Violates the rules for partial replicas. (Error 3590)
keywords: jeterr40.chm5003590
f1_keywords:
- jeterr40.chm5003590
ms.assetid: e8cb495b-cf7d-3a81-f49c-d1c8f837956e
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# Violates the rules for partial replicas. (Error 3590)

  

**Applies to:** Access 2013 | Access 2016

You cannot update a column in a table in a partial replica when another table references that column. Most likely, this is an update RI case, where the related table information does not exist in the partial replica. Make sure you follow all relationships to related tables when defining your Partial Filters.

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]