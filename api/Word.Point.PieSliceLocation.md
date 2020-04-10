---
title: Point.PieSliceLocation method (Word)
keywords: vbawd10.chm262146656
f1_keywords:
- vbawd10.chm262146656
ms.prod: word
api_name:
- Word.Point.PieSliceLocation
ms.assetid: 85687cf7-b9a8-a51d-886c-c45092cbd929
ms.date: 06/08/2017
localization_priority: Normal
---


# Point.PieSliceLocation method (Word)

Returns the vertical or horizontal position of a point on a chart item, in [points](../language/glossary/vbe-glossary.md#point), from the top or left edge of the object to the top or left edge of the chart area.


## Syntax

_expression_.**PieSliceLocation** (_loc_, _Index_)

_expression_ A variable that represents a '[Point](Word.Point.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _loc_|Required| **[XlPieSliceLocation](Word.xlpieslicelocation.md)**|Specifies a horizontal or vertical coordinate.|
| _Index_|Optional| **[XlPieSliceIndex](Word.xlpiesliceindex.md)**|Specifies which pie slice position coordinate to return. The default value is **xlOuterCenterPoint**.|

## Return value

Double


## Remarks

This property only applies to pie chart types.


## See also


[Point Object](Word.Point.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]