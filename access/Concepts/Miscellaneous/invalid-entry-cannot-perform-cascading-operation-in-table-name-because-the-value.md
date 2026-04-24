---
title: Invalid entry. Cannot perform cascading operation in table <name> because the value entered is too large for field <name>. (Error 3411)
ms.assetid: 286a606c-72c0-7dab-0dc7-0fba19d683bb
ms.date: 06/08/2019
ms.localizationpriority: medium
---


# Invalid entry. Cannot perform cascading operation in table \<name\> because the value entered is too large for field \<name\>. (Error 3411)

  

**Applies to:** Access 2013 | Access 2016

You are trying to save a value to a primary key field that is included in a relationship.

In Microsoft Access, the **Cascade Update Related Fields** option is selected for the relationship; or, in DAO code, the **dbRelationUpdateCascade** option is specified for the **Relation** object's **Attributes** property. Setting either one of these options will cause your application to attempt to update the matching field in the related table.
To save your changes to this field, the text you enter must be no longer than the field size of the related field that your application is trying to update for you. In this case, the related field has a shorter field size than the field you are updating. To avoid this problem in the future, set the **Size** property for both of the matching fields to the same value.

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](https://learn.microsoft.com/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]