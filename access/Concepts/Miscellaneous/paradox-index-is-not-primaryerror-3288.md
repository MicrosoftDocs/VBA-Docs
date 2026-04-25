---
title: Paradox index is not primary. (Error 3288)
keywords: jeterr40.chm5003288
f1_keywords:
- jeterr40.chm5003288
ms.assetid: 5866732f-b013-a777-a3ed-46a9d2ce43ea
ms.date: 06/08/2019
ms.localizationpriority: medium
---


# Paradox index is not primary. (Error 3288)

  

**Applies to:** Access 2013 | Access 2016

You are attempting to create an index on a Paradox table for which no other indexes are defined. The first index created on a Paradox table must be a primary index.

Possible solutions:


- Redefine the index as the primary index.
    
- If the index you are attempting to create is not the primary index, create the primary index first.
    

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]