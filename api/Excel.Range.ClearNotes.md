---
title: Range.ClearNotes method (Excel)
keywords: vbaxl10.chm144097
f1_keywords:
- vbaxl10.chm144097
ms.prod: excel
api_name:
- Excel.Range.ClearNotes
ms.assetid: 24017be9-d3bf-2e8a-4587-d5b0a03fdcaf
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.ClearNotes method (Excel)

Clears notes and sound notes from all the cells in the specified range.


## Syntax

_expression_. `ClearNotes`

_expression_ A variable that represents a [Range](excel.range-graph-property.md) object.


## Return value

Variant


## Example

This example clears all notes and sound notes from columns A through C on Sheet1.


```vb
Worksheets("Sheet1").Columns("A:C").ClearNotes
```


## See also


[Range Object](Excel.Range(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]