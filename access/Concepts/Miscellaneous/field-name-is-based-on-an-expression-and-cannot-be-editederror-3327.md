---
title: Field <name> is based on an expression and cannot be edited. (Error 3327)
ms.assetid: 7d7c1e1f-645e-b111-60c3-666640d8bde1
ms.date: 06/08/2019
ms.localizationpriority: medium
---


# Field \<name\> is based on an expression and cannot be edited. (Error 3327)

  

**Applies to:** Access 2013 | Access 2016

For example, if a stored query or view with a column made up of an expression was created, you would not be able to update that column. The following would return this error: CREATE VIEW VCustomer AS SELECT (FirstName & LastName) AS Test FROM Customer followed by UPDATE Test FROM VCustomer

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](https://learn.microsoft.com/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]