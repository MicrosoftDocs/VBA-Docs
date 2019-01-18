---
title: Subqueries cannot be used in the expression <expression>. (Error 3203)
keywords: jeterr40.chm5003203
f1_keywords:
- jeterr40.chm5003203
ms.prod: access
ms.assetid: 08f9c7e0-0c79-e88d-8194-ede635c49f49
ms.date: 06/08/2017
localization_priority: Normal
---


# Subqueries cannot be used in the expression <expression>. (Error 3203)

  

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

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]