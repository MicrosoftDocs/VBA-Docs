---
title: Point.PieSliceLocation method (PowerPoint)
keywords: vbapp10.chm714011
f1_keywords:
- vbapp10.chm714011
ms.prod: powerpoint
api_name:
- PowerPoint.Point.PieSliceLocation
ms.assetid: 9af5f72b-3626-9f49-09e5-6fdde51f238e
ms.date: 06/08/2017
localization_priority: Normal
---


# Point.PieSliceLocation method (PowerPoint)

Returns the vertical or horizontal position, in [points](../language/glossary/vbe-glossary.md#point), of a point on a chart item from the top or left edge of the object to the top or left edge of the chart area.


## Syntax

_expression_.**PieSliceLocation** (_loc_, _Index_)

_expression_ A variable that represents a [Point](PowerPoint.Point.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _loc_|Required|**[XlPieSliceLocation](Excel.XlPieSliceLocation.md)**|Specifies a horizontal or vertical coordinate.|
| _Index_|Optional|**[XlPieSliceIndex](Excel.XlPieSliceIndex.md)**|Specifies which pie slice position coordinate to return. The default is **xlOuterCenterPoint**.|

## Return value

Double


## Remarks

This property applies only to pie chart types.


## See also


[Point Object](PowerPoint.Point.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]