---
title: Range.Clear method (Excel)
keywords: vbaxl10.chm144094
f1_keywords:
- vbaxl10.chm144094
ms.prod: excel
api_name:
- Excel.Range.Clear
ms.assetid: 56f46ac7-8bb0-2651-8024-312c7cb7356c
ms.date: 06/08/2017
localization_priority: Priority
---


# Range.Clear method (Excel)

Clears the entire object.


## Syntax

_expression_. `Clear`

_expression_ A variable that represents a [Range](excel.range-graph-property.md) object.


## Return value

Variant


## Example

This example clears the formulas and formatting in cells A1:G37 on Sheet1.


```vb
Worksheets("Sheet1").Range("A1:G37").Clear
```


## See also


[Range Object](Excel.Range(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]