---
title: Cannot group on fields selected with '*'. (Error 3121)
keywords: jeterr40.chm5003121
f1_keywords:
- jeterr40.chm5003121
ms.prod: access
ms.assetid: fd035675-9313-f699-ab31-409063fca748
ms.date: 06/08/2019
localization_priority: Normal
---


# Cannot group on fields selected with '*'. (Error 3121)

  

**Applies to:** Access 2013 | Access 2016

You tried to execute a SELECT statement that groups or totals all fields from all tables, selected with an asterisk ( * ).

Possible cause:


- You created an SQL statement that includes an aggregate function or GROUP BY clause that refers to a field you selected with an asterisk. This error occurs, for example, if you enter the following SQL statement:
    
```sql
  SELECT * FROM Orders GROUP BY ShipVia; 

```

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]