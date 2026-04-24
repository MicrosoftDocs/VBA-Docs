---
title: No field defined -- cannot append TableDef or Index. (Error 3264)
ms.assetid: 18353c1b-c3c7-9f41-eb2a-87d732d2127a
ms.date: 06/08/2019
ms.localizationpriority: medium
---


# No field defined -- cannot append TableDef or Index. (Error 3264)

  

**Applies to:** Access 2013 | Access 2016

You cannot append a **TableDef** until you define one or more fields. Use the **CreateField** method to create fields, append them to the **Fields** collection of your **TableDef** object, and then append the **TableDef** object to the **TableDefs** collection.

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]