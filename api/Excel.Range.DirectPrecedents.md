---
title: Range.DirectPrecedents property (Excel)
keywords: vbaxl10.chm144119
f1_keywords:
- vbaxl10.chm144119
ms.prod: excel
api_name:
- Excel.Range.DirectPrecedents
ms.assetid: d7eebe51-3e4c-e902-e6a5-1617bd21ef4e
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.DirectPrecedents property (Excel)

Returns a  **[Range](Excel.Range(object).md)** object that represents the range containing all the direct precedents of a cell. This can be a multiple selection (a union of **Range** objects) if there's more than one precedent. Read-only **Range** object.


## Syntax

_expression_. `DirectPrecedents`

_expression_ A variable that represents a [Range](excel.range-graph-property.md) object.


## Remarks

The  **DirectPrecedents** property only works on the active sheet and can not trace remote references.


## Example

This example selects the direct precedents of cell A1 on Sheet1.


```vb
Worksheets("Sheet1").Activate 
Range("A1").DirectPrecedents.Select
```


## See also


[Range Object](Excel.Range(object).md)

