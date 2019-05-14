---
title: ShapeRange.Align method (Excel)
keywords: vbaxl10.chm640077
f1_keywords:
- vbaxl10.chm640077
ms.prod: excel
api_name:
- Excel.ShapeRange.Align
ms.assetid: 7a4e6442-6730-ab7d-93b5-4c091ada6b14
ms.date: 05/14/2019
localization_priority: Normal
---


# ShapeRange.Align method (Excel)

Aligns the shapes in the specified range of shapes.


## Syntax

_expression_.**Align** (_AlignCmd_, _RelativeTo_)

_expression_ A variable that represents a **[ShapeRange](Excel.shaperange.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _AlignCmd_|Required| **[MsoAlignCmd](Office.MsoAlignCmd.md)**|Specifies the way that the shapes in the specified shape range are to be aligned.|
| _RelativeTo_|Required| **[MsoTriState](Office.MsoTriState.md)**|Not used in Microsoft Excel. Must be **False**.|

## Example

This example aligns the left edges of all the shapes in the specified range in _myDocument_ with the left edge of the leftmost shape in the range.


```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes.SelectAll 
Selection.ShapeRange.Align msoAlignLefts, False
```


## See also


[ShapeRange Object](Excel.ShapeRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]