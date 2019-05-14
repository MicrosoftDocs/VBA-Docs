---
title: ShapeRange.Group method (Excel)
keywords: vbaxl10.chm640086
f1_keywords:
- vbaxl10.chm640086
ms.prod: excel
api_name:
- Excel.ShapeRange.Group
ms.assetid: f0ad9b81-42ad-0ee6-d2e2-ff2a88d47a97
ms.date: 05/14/2019
localization_priority: Normal
---


# ShapeRange.Group method (Excel)

Groups the shapes in the specified range.


## Syntax

_expression_.**Group**

_expression_ A variable that represents a **[ShapeRange](Excel.shaperange.md)** object.


## Return value

A **[Shape](Excel.Shape.md)** object that represents the grouped shape.


## Remarks

Because a group of shapes is treated as a single shape, grouping and ungrouping shapes changes the number of items in the **[Shapes](Excel.Shapes.md)** collection and changes the index numbers of items that come after the affected items in the collection.

The **[Range](Excel.Range(object).md)** object must be a single cell in the PivotTable field's data range. If you attempt to apply this method to more than one cell, it will fail (without displaying an error message).




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]