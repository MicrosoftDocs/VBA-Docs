---
title: Duplicate output destination <name>. (Error 3063)
ms.assetid: 571ce765-6e34-6860-ced2-c89733761782
ms.date: 06/08/2019
ms.localizationpriority: medium
---


# Duplicate output destination \<name\>. (Error 3063)

  

**Applies to:** Access 2013 | Access 2016

You tried to execute a query that contains more than one destination field with the same name.

Possible cause:


- You created an SQL statement that includes an INSERT INTO, SELECT...INTO, or UPDATE statement that lists the specified field name more than once.
    

Remove the duplicate fields or alias the destination field names, and try the operation again.

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](https://learn.microsoft.com/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]