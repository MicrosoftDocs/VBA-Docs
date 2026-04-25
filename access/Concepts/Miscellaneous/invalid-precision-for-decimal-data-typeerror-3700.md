---
title: Invalid precision for decimal data type. (Error 3700)
ms.assetid: 74d04222-2ed3-1552-4b56-f205633efd7f
ms.date: 06/08/2019
ms.localizationpriority: medium
---


# Invalid precision for decimal data type. (Error 3700)

  

**Applies to:** Access 2013 | Access 2016

The precision and scale for a DECIMAL data type must be between 0 and 28. For example, the following SQL statement would return this error: CREATE TABLE foo (foo DECIMAL(38,0));

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]