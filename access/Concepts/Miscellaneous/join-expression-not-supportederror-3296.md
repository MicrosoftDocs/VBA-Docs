---
title: Join expression not supported. (Error 3296)
ms.assetid: 42ae73b1-2543-1850-13a3-57ed42c54720
ms.date: 06/08/2019
ms.localizationpriority: medium
---


# Join expression not supported. (Error 3296)

  

**Applies to:** Access 2013 | Access 2016

Possible causes:



- Your SQL statement contains multiple joins in which the results of the query can differ, depending on the order in which the joins are performed. You may want to create a separate query to perform the first join, and then include that query in your SQL statement.
    
- The ON statement in your JOIN operation is incomplete or contains too many tables. You may want to put your ON expression in a WHERE clause.
    

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](https://learn.microsoft.com/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]