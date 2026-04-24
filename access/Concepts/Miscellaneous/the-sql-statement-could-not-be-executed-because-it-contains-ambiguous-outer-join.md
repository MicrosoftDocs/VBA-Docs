---
title: The SQL statement could not be executed because it contains ambiguous outer joins. To force one of the joins to be performed first, create a separate query that performs the first join and then include that query in your SQL statement. (Error 3258)
keywords: jeterr40.chm5003258
f1_keywords:
- jeterr40.chm5003258
ms.assetid: 17515e13-d6d8-8a1e-ee6c-ff2af543da0f
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# The SQL statement could not be executed because it contains ambiguous outer joins. To force one of the joins to be performed first, create a separate query that performs the first join and then include that query in your SQL statement. (Error 3258)

  

**Applies to:** Access 2013 | Access 2016

You tried to execute an SQL statement that contains multiple joins; the results of the query can differ depending on the order in which the joins are performed. For example, this error can occur if you execute the following SQL statement:




```sql
SELECT * FROM Customers, Orders, [Order Details],
Customers LEFT JOIN Orders 
ON Customers.CustomerID = Orders.CustomerID, 
Orders INNER JOIN [Order Details] 
ON Orders.OrderID = [Order Details].OrderID;

```

Executing this statement produces an error because the order of the joins is ambiguous. To force one of the joins to be performed first, create a separate query that performs the first join and then include that query in your SQL statement. The following queries illustrate how you might construct the preceding query so that the INNER JOIN operation is performed before the LEFT JOIN and RIGHT JOIN operation:
Query1



```sql
SELECT * FROM Orders, [Order Details],
Orders INNER JOIN [Order Details]
ON Orders. OrderID = [Order Details].OrderID;
```

Query2



```sql
SELECT * FROM Customers, Query1,
Customers LEFT JOIN Query1 
ON Customers.CustomerID = Orders.CustomerID;
```

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](https://learn.microsoft.com/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
