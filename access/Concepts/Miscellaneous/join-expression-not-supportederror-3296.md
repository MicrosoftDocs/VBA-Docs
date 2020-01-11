---
title: Join expression not supported. (Error 3296)
ms.prod: access
ms.assetid: 42ae73b1-2543-1850-13a3-57ed42c54720
ms.date: 06/08/2017
localization_priority: Normal
---


# Join expression not supported. (Error 3296)

  

**Applies to:** Access 2013 | Access 2016

Possible causes:



- Your SQL statement contains multiple joins in which the results of the query can differ, depending on the order in which the joins are performed. You may want to create a separate query to perform the first join, and then include that query in your SQL statement.
    
- The ON statement in your JOIN operation is incomplete or contains too many tables. You may want to put your ON expression in a WHERE clause.
    

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]