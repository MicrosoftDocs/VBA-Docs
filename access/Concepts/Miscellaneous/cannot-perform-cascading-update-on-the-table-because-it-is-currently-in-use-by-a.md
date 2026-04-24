---
title: Cannot perform cascading update on the table because it is currently in use by another user. (Error 3412)
keywords: jeterr40.chm5003412
f1_keywords:
- jeterr40.chm5003412
ms.assetid: 0718b58e-5553-8c08-ea85-83f97eb88998
ms.date: 06/08/2019
ms.localizationpriority: medium
---


# Cannot perform cascading update on the table because it is currently in use by another user. (Error 3412)

  

**Applies to:** Access 2013 | Access 2016

You are trying to save a value into a primary key field that is included in a relationship.

In Microsoft Access, the **Cascade Update Related Fields** option is selected for the relationship, or in DAO code, the **dbRelationUpdateCascade** option is specified for the **Relation** object's **Attributes** property. Your application is attempting to update the matching field in the related table.
The matching field cannot be updated because of a locking conflict with another user. To save the record, you must wait until the related table is no longer in use.


## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](https://learn.microsoft.com/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]