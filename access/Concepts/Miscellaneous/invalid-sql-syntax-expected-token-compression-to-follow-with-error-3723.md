---
title: "Invalid SQL Syntax: expected token: COMPRESSION to follow WITH (Error 3723)"
ms.prod: access
ms.assetid: 6d63cc77-dbcf-302d-6957-1439f18dceeb
ms.date: 06/08/2019
localization_priority: Normal
---


# Invalid SQL Syntax: expected token: COMPRESSION to follow WITH (Error 3723)

  

**Applies to:** Access 2013 | Access 2016

This error occurs when using CREATE TABLE or ALTER TABLE ALTER COLUMN syntax. It occurs when referencing one of the NATIONAL CHARACTER synonyms for a column type and not using the COMPRESSION keyword following the WITH keyword. The following is a valid SQL statement: CREATE TABLE foo (foo NCHAR WITH COMP);. The following SQL statement would return the error: CREATE TABLE foo (foo NCHAR WITH);.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]