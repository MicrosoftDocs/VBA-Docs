---
title: Range.IndentLevel property (Excel)
keywords: vbaxl10.chm144147
f1_keywords:
- vbaxl10.chm144147
ms.prod: excel
api_name:
- Excel.Range.IndentLevel
ms.assetid: f4d5af31-904a-27eb-fb2d-e5ae38a7ebb9
ms.date: 05/11/2019
localization_priority: Normal
---


# Range.IndentLevel property (Excel)

Returns or sets a **Variant** value that represents the indent level for the cell or range. Can be an integer from 0 to 15.


## Syntax

_expression_.**IndentLevel**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Remarks

Using this property to set the indent level to a number less than 0 (zero) or greater than 15 causes an error.


## Example

This example increases the indent level to 15 in cell A10.

```vb
With Range("A10") 
 .IndentLevel = 15 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
