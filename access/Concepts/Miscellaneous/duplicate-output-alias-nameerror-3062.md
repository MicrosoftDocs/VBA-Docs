---
title: Duplicate output alias <name>. (Error 3062)
keywords: jeterr40.chm5003062
f1_keywords:
- jeterr40.chm5003062
ms.prod: access
ms.assetid: e0157e7c-d854-4a9a-b5ba-22afa0944cbc
ms.date: 06/08/2017
localization_priority: Normal
---


# Duplicate output alias <name>. (Error 3062)

  

**Applies to:** Access 2013 | Access 2016

You tried to execute an SQL statement that has more than one alias with the same name. The following statement, for example, would produce this error:




```sql
SELECT LastName AS Name, FirstName AS Name FROM Employees;

```

Rename one or more of the aliases, and then try the operation again.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
