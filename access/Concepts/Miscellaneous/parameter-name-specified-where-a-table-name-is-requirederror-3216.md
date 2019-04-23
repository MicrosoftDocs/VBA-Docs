---
title: Parameter <name> specified where a table name is required. (Error 3216)
keywords: jeterr40.chm5003216
f1_keywords:
- jeterr40.chm5003216
ms.prod: access
ms.assetid: baf55d2f-1f19-bb4c-1fdd-339ea4024638
ms.date: 06/08/2017
localization_priority: Normal
---


# Parameter <name> specified where a table name is required. (Error 3216)

  

**Applies to:** Access 2013 | Access 2016

You created a parameter query that specifies an invalid parameter type. The following example produces this error.




```sql
PARAMETERS Param1 Text; 

INSERT INTO Param1 
SELECT * 
FROM Customers; 

```

 `Param1` is a text parameter, but the INSERT INTO statement expects a table name.
Change the parameter type from Text to TableID, and then try the operation again.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]