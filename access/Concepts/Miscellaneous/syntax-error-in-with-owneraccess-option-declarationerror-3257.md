---
title: Syntax error in WITH OWNERACCESS OPTION declaration. (Error 3257)
keywords: jeterr40.chm5003257
f1_keywords:
- jeterr40.chm5003257
ms.prod: access
ms.assetid: a1b4ae18-4efa-d79a-ffec-4ec705a0236b
ms.date: 06/08/2017
localization_priority: Normal
---


# Syntax error in WITH OWNERACCESS OPTION declaration. (Error 3257)

  

**Applies to:** Access 2013 | Access 2016

Possible causes:



- The WITH OWNERACCESS OPTION declaration is incomplete or includes a space between OWNER and ACCESS.
    
- The declaration appears in an unexpected and disallowed position in the SQL statement. For example:
    
```sql
  SELECT * WITH OWNERACCESS OPTION FROM [My Table]; 

```


    The WITH OWNERACCESS OPTION declaration should appear at the end of the SQL statement, usually after the ORDER BY clause, if present:
    


```sql
  SELECT * FROM [My Table] WITH OWNERACCESS OPTION;
```

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]