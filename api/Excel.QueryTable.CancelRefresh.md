---
title: QueryTable.CancelRefresh method (Excel)
keywords: vbaxl10.chm518082
f1_keywords:
- vbaxl10.chm518082
ms.prod: excel
api_name:
- Excel.QueryTable.CancelRefresh
ms.assetid: be9491bd-9b42-4b88-ddb9-554cf431e779
ms.date: 05/03/2019
localization_priority: Normal
---


# QueryTable.CancelRefresh method (Excel)

Cancels all background queries for the specified query table. Use the **[Refreshing](Excel.QueryTable.Refreshing.md)** property to determine whether a background query is currently in progress.


## Syntax

_expression_.**CancelRefresh**

_expression_ A variable that represents a **[QueryTable](Excel.QueryTable.md)** object.


## Example

This example cancels a query table refresh operation.

```vb
With Worksheets(1).QueryTables(1) 
 If .Refreshing Then .CancelRefresh 
End With 

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]