---
title: Duplicate output destination <name>. (Error 3063)
ms.prod: access
ms.assetid: 571ce765-6e34-6860-ced2-c89733761782
ms.date: 06/08/2019
localization_priority: Normal
---


# Duplicate output destination <name>. (Error 3063)

  

**Applies to:** Access 2013 | Access 2016

You tried to execute a query that contains more than one destination field with the same name.

Possible cause:


- You created an SQL statement that includes an INSERT INTO, SELECT...INTO, or UPDATE statement that lists the specified field name more than once.
    

Remove the duplicate fields or alias the destination field names, and try the operation again.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]