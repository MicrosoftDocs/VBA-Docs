---
title: ChartObject.ShapeRange property (Excel)
keywords: vbaxl10.chm494097
f1_keywords:
- vbaxl10.chm494097
ms.prod: excel
api_name:
- Excel.ChartObject.ShapeRange
ms.assetid: 12ad4077-1687-2bb9-41cf-fd8f1e02adc0
ms.date: 04/20/2019
localization_priority: Normal
---


# ChartObject.ShapeRange property (Excel)

Returns a **[ShapeRange](Excel.ShapeRange.md)** object that represents the specified object or objects. Read-only.


## Syntax

_expression_.**ShapeRange**

_expression_ An expression that returns a **[ChartObject](Excel.ChartObject.md)** object.


## Example

This example creates a shape range that represents the embedded charts on worksheet one.

```vb
Set sr = Worksheets(1).ChartObjects.ShapeRange
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]