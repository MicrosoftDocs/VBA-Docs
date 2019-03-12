---
title: Build SQL statements that include variables and controls
ms.prod: access
ms.assetid: e902199f-ed00-e885-3671-0705aa2b058a
ms.date: 06/08/2017
localization_priority: Normal
---


# Build SQL statements that include variables and controls

When working with Data Access Objects (DAO) or ActiveX Data Objects (ADO), you may need to construct an SQL statement in code. This is sometimes referred to as taking your SQL code "inline." 

For example, if you are creating a new [QueryDef](../../../api/overview/Access.md) object, you must set its [SQL](../../../api/overview/Access.md) property to a valid SQL string. But if you are using an ADO [Recordset](../../../api/overview/Access.md) object, you must set its [Source](../../../api/overview/Access.md) property to a valid SQL string.

To construct an SQL statement, create a query in the query design grid, switch to SQL view, and copy and paste the corresponding SQL statement into your code.

Often a query must be based on values that the user supplies, or values that change in different situations. If this is the case, you need to include variables or control values in your query. The Access database engine processes all SQL statements, but not variables or controls. Therefore, you must construct your SQL statement so that Access first determines these values and then concatenates them into the SQL statement that is passed to the Access database engine.


## Build SQL statements with DAO

The following example shows how to create a **QueryDef** object with a simple SQL statement. This query returns all orders from an Orders table that were placed after March 31, 2006.


```vb
Public Sub GetOrders() 
 
   Dim dbs As DAO.Database 
   Dim qdf As DAO.QueryDef 
   Dim strSQL As String 
 
   Set dbs = CurrentDb 
   strSQL = "SELECT * FROM Orders WHERE OrderDate >#3-31-2006#;" 
   Set qdf = dbs.CreateQueryDef("SecondQuarter", strSQL) 
 
End Sub
```

The next example creates the same **QueryDef** object by using a value stored in a variable. Be aware that the number signs (#) that denote the date values must be included in the string so that they are concatenated with the date value.

```vb
Dim dbs As Database, qdf As QueryDef, strSQL As String 
Dim dteStart As Date 
dteStart = #3-31-2006# 
Set dbs = CurrentDb 
strSQL = "SELECT * FROM Orders WHERE OrderDate" _ 
    & "> #" & dteStart & "#;" 
Set qdf = dbs.CreateQueryDef("SecondQuarter", strSQL)
```

The following example creates a **QueryDef** object by using a value in a control called OrderDate on an Orders form. Be aware that you provide the full reference to the control, and that you include the number signs (#) that denote the date within the string.

```vb
Dim dbs As Database, qdf As QueryDef, strSQL As String 
Set dbs = CurrentDb 
strSQL = "SELECT * FROM Orders WHERE OrderDate" _ 
    & "> #" & Forms!Orders!OrderDate & "#;" 
Set qdf = dbs.CreateQueryDef("SecondQuarter", strSQL)
```


## Build SQL statements with ADO

In this section, you will build the same statements as in the previous section, but this time using ADO as the data access method.

The following code example shows how to create a **QueryDef** object with a simple SQL statement. This query returns all orders from an Orders table that were placed after March 31, 2006.

```vb
Dim dbs As Database, qdf As QueryDef, strSQL As String 
Set dbs = CurrentDb 
strSQL = "SELECT * FROM Orders WHERE OrderDate >#3-31-2006#;" 
Set qdf = dbs.CreateQueryDef("SecondQuarter", strSQL)
```

The next example creates the same **QueryDef** object by using a value stored in a variable. Be aware that the number signs (#) that denote the date values must be included in the string so that they are concatenated with the date value.

```vb
Dim dbs As Database, qdf As QueryDef, strSQL As String 
Dim dteStart As Date 
dteStart = #3-31-2006# 
Set dbs = CurrentDb 
strSQL = "SELECT * FROM Orders WHERE OrderDate" _ 
    & "> #" & dteStart & "#;" 
Set qdf = dbs.CreateQueryDef("SecondQuarter", strSQL)
```

The following code example creates a **QueryDef** object by using a value in a control called OrderDate on an Orders form. Be aware that it provides the full reference to the control, and that it includes the number signs that denote the date within the string.

```vb
Dim dbs As Database, qdf As QueryDef, strSQL As String 
Set dbs = CurrentDb 
strSQL = "SELECT * FROM Orders WHERE OrderDate" _ 
    & "> #" & Forms!Orders!OrderDate & "#;" 
Set qdf = dbs.CreateQueryDef("SecondQuarter", strSQL)
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
