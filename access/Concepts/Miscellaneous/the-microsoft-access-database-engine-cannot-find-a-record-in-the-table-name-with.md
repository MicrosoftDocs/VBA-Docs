---
title: The Microsoft Access database engine cannot find a record in the table <name> with key matching field(s) <name>. (Error 3101)
keywords: jeterr40.chm5003101
f1_keywords:
- jeterr40.chm5003101
ms.assetid: 6842aab2-7f8d-354f-9690-dfda2d2380a5
ms.date: 06/08/2017
ms.localizationpriority: high
---


# The Microsoft Access database engine cannot find a record in the table \<name\> with key matching field(s) \<name\>. (Error 3101)

  

**Applies to:** Access 2013 | Access 2016

In a one-to-many relationship, you entered data on the "many" side for which there is no matching record on the "one" side. For example, this error occurs if you join a Customers table and Orders table on a CustomerID field, and then add an order using a CustomerID that does not exist in the Customers table.

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
