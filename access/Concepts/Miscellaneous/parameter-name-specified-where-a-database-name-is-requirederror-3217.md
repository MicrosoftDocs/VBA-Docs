---
title: Parameter <name> specified where a database name is required. (Error 3217)
keywords: jeterr40.chm5003217
f1_keywords:
- jeterr40.chm5003217
ms.prod: access
ms.assetid: d6d700f2-5df5-5d26-a6ee-706ca4c1a12a
ms.date: 06/08/2017
localization_priority: Normal
---


# Parameter <name> specified where a database name is required. (Error 3217)

  

**Applies to:** Access 2013 | Access 2016

You created a parameter query that specifies an invalid parameter type. The following example produces this error:




```sql
PARAMETERS Param1 Text;

SELECT CustomerID
FROM Customers IN Param1;
```

 `Param1` is a text parameter, but the FROM clause requires a database parameter.
Change the parameter type from Text to Database, and then try the operation again.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]