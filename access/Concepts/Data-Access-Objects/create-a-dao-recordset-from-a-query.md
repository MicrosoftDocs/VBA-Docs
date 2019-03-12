---
title: Create a DAO Recordset from a query
ms.prod: access
ms.assetid: d84870d4-58e4-9d48-9951-72d928929002
ms.date: 09/21/2018
localization_priority: Normal
---


# Create a DAO Recordset from a query

You can create a **[Recordset](../../../api/overview/Access.md)** object based on a stored select query. In the following code example, `Current Product List` is an existing select query stored in the current database.


```vb
Dim dbsNorthwind As DAO.Database 
Dim rstProducts As DAO.Recordset 
 
Set dbsNorthwind = CurrentDb 
Set rstProducts = dbsNorthwind.OpenRecordset("Current Product List") 

```


If a stored select query does not already exist, the **[OpenRecordset](../../../api/overview/Access.md)** method also accepts an SQL string instead of the name of a query. The previous example can be rewritten as follows.

```vb
Dim dbsNorthwind As DAO.Database 
Dim rstProducts As DAO.Recordset 
Dim strSQL As String 
 
Set dbsNorthwind = CurrentDb 
strSQL = "SELECT * FROM Products WHERE Discontinued = No " & _ 
         "ORDER BY ProductName" 
Set rstProducts = dbsNorthwind.OpenRecordset(strSQL) 

```

The disadvantage of this approach is that the query string must be compiled each time it runs, whereas the stored query is compiled the first time it is saved, which usually results in slightly better performance.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
