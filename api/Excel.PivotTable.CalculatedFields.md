---
title: PivotTable.CalculatedFields method (Excel)
keywords: vbaxl10.chm235103
f1_keywords:
- vbaxl10.chm235103
ms.prod: excel
api_name:
- Excel.PivotTable.CalculatedFields
ms.assetid: 8f09c79d-48e7-0c75-8db2-2201fcdcc974
ms.date: 05/08/2019
localization_priority: Normal
---


# PivotTable.CalculatedFields method (Excel)

Returns a **[CalculatedFields](Excel.CalculatedFields.md)** collection that represents all the calculated fields in the specified PivotTable report. Read-only.


## Syntax

_expression_.**CalculatedFields**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Return value

CalculatedFields


## Example

This example prevents the calculated fields from being dragged to the row position.

```vb
For Each fld in _ 
 Worksheets(1).PivotTables("Pivot1") _ 
 .CalculatedFields 
 fld.DragToRow = False 
Next
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]