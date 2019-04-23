---
title: Parameters object (Excel)
keywords: vbaxl10.chm524072
f1_keywords:
- vbaxl10.chm524072
ms.prod: excel
api_name:
- Excel.Parameters
ms.assetid: d67147f1-d587-a9e4-ed8e-8a1140e8a868
ms.date: 03/30/2019
localization_priority: Normal
---


# Parameters object (Excel)

A collection of **[Parameter](Excel.Parameter.md)** objects for the specified query table.


## Remarks

Each **Parameter** object represents a single query parameter. Every query table contains a **Parameters** collection, but the collection is empty unless the query table is using a parameter query.

You cannot use the **Add** method on a URL connection query table. For URL connection query tables, Microsoft Excel creates the parameters based on the **[Connection](Excel.QueryTable.Connection.md)** and **[PostText](Excel.QueryTable.PostText.md)** properties.


## Example

Use the **[Parameters](excel.querytable.parameters.md)** property of the **QueryTable** object to return the **Parameters** collection. 

The following example displays the number of parameters in query table one.

```vb
MsgBox Workbooks(1).ActiveSheet.QueryTables(1).Parameters.Count
```

<br/>

Use the **Add** method to create a new parameter for a query table. The following example changes the SQL statement for query table one. The clause "(city=?)" indicates that the query is a parameter query, and the value of city is set to the constant Oakland.

```vb
Set qt = Sheets("sheet1").QueryTables(1) 
qt.Sql = "SELECT * FROM authors WHERE (city=?)" 
Set param1 = qt.Parameters.Add("City Parameter", _ 
 xlParamTypeVarChar) 
param1.SetParam xlConstant, "Oakland" 
qt.Refresh
```

## Methods

- [Add](Excel.Parameters.Add.md)
- [Delete](Excel.Parameters.Delete.md)
- [Item](Excel.Parameters.Item.md)

## Properties

- [Application](Excel.Parameters.Application.md)
- [Count](Excel.Parameters.Count.md)
- [Creator](Excel.Parameters.Creator.md)
- [Parent](Excel.Parameters.Parent.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]