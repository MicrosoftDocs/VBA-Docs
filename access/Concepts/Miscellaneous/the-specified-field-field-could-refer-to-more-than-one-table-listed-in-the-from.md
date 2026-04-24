---
title: The specified field <field> could refer to more than one table listed in the FROM clause of your SQL statement. (Error 3079)
keywords: jeterr40.chm5003079
f1_keywords:
- jeterr40.chm5003079
ms.assetid: 5dcb65e3-3f8c-f16c-5380-1d665283aa7a
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# The specified field \<field\> could refer to more than one table listed in the FROM clause of your SQL statement. (Error 3079)

  

**Applies to:** Access 2013 | Access 2016

The specified field reference could refer to more than one table listed in the FROM clause of your SQL statement. In the following example, the OrderID field exists in both the Orders and Order Details tables:




```sql
SELECT OrderID 
FROM Orders, [Order Details];
```

Because the statement does not specify which table OrderID belongs to, it produces this error. To complete this operation, fully qualify the field reference by adding a table name. For example:



```sql
SELECT Orders.OrderID 
FROM Orders, [Order Details];
```

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
