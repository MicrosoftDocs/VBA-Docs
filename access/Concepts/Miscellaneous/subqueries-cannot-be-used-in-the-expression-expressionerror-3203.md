---
title: Subqueries cannot be used in the expression <expression>. (Error 3203)
keywords: jeterr40.chm5003203
f1_keywords:
- jeterr40.chm5003203
ms.assetid: 08f9c7e0-0c79-e88d-8194-ede635c49f49
ms.date: 06/08/2019
ms.localizationpriority: medium
---

# Subqueries cannot be used in the expression \<expression\>. (Error 3203)


**Applies to:** Access 2013 | Access 2016

The specified expression contains a subquery or other expression that functions as a subquery.

Possible cause:

- You used a SELECT statement that includes an aggregate function that evaluates another aggregate function. This error occurs, for example, if you execute the following statement: 
    
```sql
  TRANSFORM Sum([OrderAmount]) AS Sum1 
SELECT Sum([Sum1]) AS Sum2, OrderID, Sum([Sum2]) AS Expr1 
FROM Orders 
GROUP BY OrderID 
PIVOT ShipName;
```

The < _expression_ > parameter of the alert would contain the expression `Sum([Sum2])` from the SELECT clause, because this references an alias used in the same SELECT statement and acts as a subquery against `Sum([Sum1])`.
    
## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](https://learn.microsoft.com/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
