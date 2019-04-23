---
title: ChartObjects.ShapeRange property (Excel)
keywords: vbaxl10.chm497095
f1_keywords:
- vbaxl10.chm497095
ms.prod: excel
api_name:
- Excel.ChartObjects.ShapeRange
ms.assetid: 4813fce5-ad3f-861c-d6dc-63fb617ed4da
ms.date: 04/20/2019
localization_priority: Normal
---


# ChartObjects.ShapeRange property (Excel)

Returns a **[ShapeRange](Excel.ShapeRange.md)** object that represents the specified object or objects. Read-only.


## Syntax

_expression_.**ShapeRange**

_expression_ An expression that returns a **[ChartObjects](Excel.ChartObjects.md)** object.


## Example

This example creates a shape range that represents the embedded charts on worksheet one.

```vb
Set sr = Worksheets(1).ChartObjects.ShapeRange
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]