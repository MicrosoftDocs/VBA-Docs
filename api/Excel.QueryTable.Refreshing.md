---
title: QueryTable.Refreshing property (Excel)
keywords: vbaxl10.chm518079
f1_keywords:
- vbaxl10.chm518079
ms.prod: excel
api_name:
- Excel.QueryTable.Refreshing
ms.assetid: 7b89fbec-3365-5b23-1b21-da3b0145d9bc
ms.date: 05/03/2019
localization_priority: Normal
---


# QueryTable.Refreshing property (Excel)

**True** if there is a background query in progress for the specified query table. Read-only **Boolean**.


## Syntax

_expression_.**Refreshing**

_expression_ A variable that represents a **[QueryTable](Excel.QueryTable.md)** object.


## Remarks

Use the **[CancelRefresh](Excel.QueryTable.CancelRefresh.md)** method to cancel background queries.

If you import data by using the user interface, data from a web query or a text query is imported as a **QueryTable** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data by using the object model, data from a web query or a text query must be imported as a **QueryTable**, while all other external data can be imported as either a **ListObject** or a **QueryTable**.

You can use the **[QueryTable](Excel.ListObject.QueryTable.md)** property of the **ListObject** to access the **Refreshing** property.


## Example

This example displays a message box if there is a background query in progress for query table one.

```vb
With Worksheets(1).QueryTables(1) 
 If .Refreshing Then 
 MsgBox "Query is currently refreshing: please wait" 
 Else 
 .Refresh BackgroundQuery := False 
 .ResultRange.Select 
 End If 
End With 

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
