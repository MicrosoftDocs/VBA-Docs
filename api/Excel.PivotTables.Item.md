---
title: PivotTables.Item method (Excel)
keywords: vbaxl10.chm238074
f1_keywords:
- vbaxl10.chm238074
ms.prod: excel
api_name:
- Excel.PivotTables.Item
ms.assetid: 1bdc8558-ec67-2823-fd02-ecd5ae4ecee6
ms.date: 05/07/2019
localization_priority: Normal
---


# PivotTables.Item method (Excel)

Returns a single object from a collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[PivotTables](Excel.PivotTables.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number for the object.|

## Return value

A **[PivotTable](Excel.PivotTable.md)** object contained by the collection.


## Remarks

The text name of the object is the value of the **[Name](Excel.PivotTable.Name.md)** and **[Value](Excel.PivotTable.Value.md)** properties.


## Example

This example makes the Year field a row field in the first PivotTable report on Sheet3.

```vb
Worksheets("sheet3").PivotTables.Item(1) _ 
 .PivotFields("year").Orientation = xlRowField
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]