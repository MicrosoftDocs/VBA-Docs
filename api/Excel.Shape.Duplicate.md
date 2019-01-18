---
title: Shape.Duplicate method (Excel)
keywords: vbaxl10.chm636076
f1_keywords:
- vbaxl10.chm636076
ms.prod: excel
api_name:
- Excel.Shape.Duplicate
ms.assetid: d4e32396-5d9e-2e10-3111-e0ceda52643e
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.Duplicate method (Excel)

Duplicates the object and returns a reference to the new copy.


## Syntax

_expression_. `Duplicate`

_expression_ A variable that represents a [Shape](./Excel.Shape.md) object.


## Return value

Shape


## Example

This example duplicates embedded chart one on Sheet1 and then selects the copy.


```vb
Set dChart = Worksheets("Sheet1").ChartObjects(1).Duplicate 
dChart.Select
```


## See also


[Shape Object](Excel.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]