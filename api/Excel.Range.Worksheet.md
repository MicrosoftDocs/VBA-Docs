---
title: Range.Worksheet property (Excel)
keywords: vbaxl10.chm144220
f1_keywords:
- vbaxl10.chm144220
ms.prod: excel
api_name:
- Excel.Range.Worksheet
ms.assetid: af38bdde-d523-a4cd-929e-1f67464b2593
ms.date: 06/08/2017
localization_priority: Priority
---


# Range.Worksheet property (Excel)

Returns a  **[Worksheet](Excel.Worksheet.md)** object that represents the worksheet containing the specified range. Read-only.


## Syntax

_expression_. `Worksheet`

_expression_ A variable that represents a [Range](excel.range-graph-property.md) object.


## Example

This example displays the name of the worksheet that contains the active cell. The example must be run from a worksheet.


```vb
MsgBox ActiveCell.Worksheet.Name
```

This example displays the name of the worksheet that contains the range named "testRange."




```vb
MsgBox Range("testRange").Worksheet.Name
```


## See also


[Range Object](Excel.Range(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]