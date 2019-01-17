---
title: Range.UnMerge method (Excel)
keywords: vbaxl10.chm144159
f1_keywords:
- vbaxl10.chm144159
ms.prod: excel
api_name:
- Excel.Range.UnMerge
ms.assetid: dfc49876-29b0-0b61-fe18-3953438f7452
ms.date: 06/08/2017
localization_priority: Priority
---


# Range.UnMerge method (Excel)

Separates a merged area into individual cells.


## Syntax

_expression_. `UnMerge`

_expression_ A variable that represents a [Range](excel.range-graph-property.md) object.


## Example

This example separates the merged range that contains cell A3.


```vb
With Range("a3") 
 If .MergeCells Then 
 .MergeArea.UnMerge 
 Else 
 MsgBox "not merged" 
 End If 
End With
```


## See also


[Range Object](Excel.Range(object).md)

