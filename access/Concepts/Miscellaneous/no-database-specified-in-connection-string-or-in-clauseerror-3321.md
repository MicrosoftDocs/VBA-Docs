---
title: No database specified in connection string or IN clause. (Error 3321)
ms.assetid: dd1601c0-9998-4ae0-21a0-ec283dc8f0cd
ms.date: 06/08/2019
ms.localizationpriority: medium
---


# No database specified in connection string or IN clause. (Error 3321)

  

**Applies to:** Access 2013 | Access 2016

The Microsoft Access database engine is unable to connect to an external ISAM database because you have not specified the name of the database to connect to. When connecting to an external source of data, you must specify a database name.

Possible causes:


- The connection string in the FROM clause of the SELECT statement is missing the parameter, DATABASE=.
    
- The IN clause of the SELECT statement includes a database type argument (indicating to select the data from an external database) but it is missing a database name argument.
    

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]