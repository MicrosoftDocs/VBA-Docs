---
title: QueryTable.FetchedRowOverflow property (Excel)
keywords: vbaxl10.chm518080
f1_keywords:
- vbaxl10.chm518080
ms.prod: excel
api_name:
- Excel.QueryTable.FetchedRowOverflow
ms.assetid: 386aaf06-27d4-bfa1-cf5e-ac8c8bddef44
ms.date: 05/03/2019
localization_priority: Normal
---


# QueryTable.FetchedRowOverflow property (Excel)

**True** if the number of rows returned by the last use of the **[Refresh](Excel.QueryTable.Refresh.md)** method is greater than the number of rows available on the worksheet. Read-only **Boolean**.


## Syntax

_expression_.**FetchedRowOverflow**

_expression_ A variable that represents a **[QueryTable](Excel.QueryTable.md)** object.


## Remarks

If you import data by using the user interface, data from a web query or a text query is imported as a **QueryTable** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data by using the object model, data from a web query or a text query must be imported as a **QueryTable**, while all other external data can be imported as either a **ListObject** or a **QueryTable**.

You can use the **[QueryTable](Excel.ListObject.QueryTable.md)** property of the **ListObject** to access the **FetchedRowOverflow** property.


## Example

This example refreshes query table one. If the number of rows returned by the query exceeds the number of rows available on the worksheet, an error message is displayed.

```vb
With Worksheets(1).QueryTables(1) 
 .Refresh 
 If .FetchedRowOverflow Then 
 MsgBox "Query too large: please redefine." 
 End If 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]