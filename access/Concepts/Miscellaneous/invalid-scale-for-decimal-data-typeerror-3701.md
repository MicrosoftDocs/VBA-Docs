---
title: Invalid scale for decimal data type. (Error 3701)
keywords: jeterr40.chm5003701
f1_keywords:
- jeterr40.chm5003701
ms.assetid: 2c839f39-d3ab-053a-d7b0-1bcde43232d4
ms.date: 06/08/2019
ms.localizationpriority: medium
---


# Invalid scale for decimal data type. (Error 3701)

  

**Applies to:** Access 2013 | Access 2016

The scale of a DECIMAL data type must always be less or equal to the precision. For example, the following SQL statement would return this error: CREATE TABLE foo (foo DECIMAL(10,12));

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](https://learn.microsoft.com/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]