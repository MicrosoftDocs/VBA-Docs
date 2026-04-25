---
title: The cascading options for the new reference conflict with existing reference <name>. (Error 3707)
keywords: jeterr40.chm5003707
f1_keywords:
- jeterr40.chm5003707
ms.assetid: 4f45ecee-ac02-d26b-7ae5-eac9a75df83e
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# The cascading options for the new reference conflict with existing reference \<name\>. (Error 3707)

  

**Applies to:** Access 2013 | Access 2016

This error occurs if a CASCADE action is defined on a column that already has another type of CASCADE action. For example, if CASCADE DELETE is already specified, the user will be prevented from trying to add CASCADE UPDATE. To apply the desired CASCADE action, the original CONSTRAINT must be dropped. This can be done with the ALTER TABLE ALTER COLUMN syntax or with the DROP CONSTRAINT syntax.

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]