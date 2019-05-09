---
title: Range.Dependents property (Excel)
keywords: vbaxl10.chm144116
f1_keywords:
- vbaxl10.chm144116
ms.prod: excel
api_name:
- Excel.Range.Dependents
ms.assetid: 47813412-306a-0f99-3ca5-d354b16af468
ms.date: 05/10/2019
localization_priority: Normal
---


# Range.Dependents property (Excel)

Returns a **Range** object that represents the range containing all the dependents of a cell. This can be a multiple selection (a union of **Range** objects) if there's more than one dependent. Read-only **Range** object.


## Syntax

_expression_.**Dependents**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Remarks

The **Dependents** property only works on the active sheet and cannot trace remote references.


## Example

This example selects the dependents of cell A1 on Sheet1.

```vb
Worksheets("Sheet1").Activate 
Range("A1").Dependents.Select
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]