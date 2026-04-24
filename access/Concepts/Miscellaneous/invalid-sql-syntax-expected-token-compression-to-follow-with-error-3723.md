---
title: "Invalid SQL Syntax: expected token: COMPRESSION to follow WITH (Error 3723)"
ms.assetid: 6d63cc77-dbcf-302d-6957-1439f18dceeb
ms.date: 06/08/2019
ms.localizationpriority: medium
---


# Invalid SQL Syntax: expected token: COMPRESSION to follow WITH (Error 3723)

  

**Applies to:** Access 2013 | Access 2016

This error occurs when using CREATE TABLE or ALTER TABLE ALTER COLUMN syntax. It occurs when referencing one of the NATIONAL CHARACTER synonyms for a column type and not using the COMPRESSION keyword following the WITH keyword. The following is a valid SQL statement: CREATE TABLE foo (foo NCHAR WITH COMP);. The following SQL statement would return the error: CREATE TABLE foo (foo NCHAR WITH);.

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](https://learn.microsoft.com/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]