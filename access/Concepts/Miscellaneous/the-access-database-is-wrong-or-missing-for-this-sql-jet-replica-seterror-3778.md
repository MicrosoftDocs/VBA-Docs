---
title: The Access database is wrong or missing for this SQL/Jet replica set. (Error 3778)
keywords: jeterr40.chm5003778
f1_keywords:
- jeterr40.chm5003778
ms.prod: access
ms.assetid: 21c104f6-06cb-f3b5-aa9e-098dca92c223
ms.date: 06/08/2017
localization_priority: Normal
---


# The Access database is wrong or missing for this SQL/Jet replica set. (Error 3778)

  

**Applies to:** Access 2013 | Access 2016

The Microsoft Access database specified is wrong or missing because:



- The data source path supplied in the link server on the SQL Server points to an invalid replica for this publication. The replica may have been replaced, damaged, or not created properly.
    
- The hub row is missing or is not valid.
    
- The database specified is not replicable.
    

The solution is to re-initialize your Jet Subscriber using the Re-Initialize tools on the SQL Server.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]