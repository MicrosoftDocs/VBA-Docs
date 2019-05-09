---
title: Point.PieSliceLocation method (Excel)
keywords: vbaxl10.chm576109
f1_keywords:
- vbaxl10.chm576109
ms.prod: excel
api_name:
- Excel.Point.PieSliceLocation
ms.assetid: 90a318d4-0ad2-d326-c26b-3c965b1ffe43
ms.date: 05/09/2019
localization_priority: Normal
---


# Point.PieSliceLocation method (Excel)

Returns the vertical or horizontal position of a point on a chart item, in [points](../language/glossary/vbe-glossary.md#point), from the top or left edge of the object to the top or left edge of the chart area.


## Syntax

_expression_.**PieSliceLocation** (_loc_, _Index_)

_expression_ A variable that represents a **[Point](Excel.Point(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _loc_|Required| **[XlPieSliceLocation](Excel.XlPieSliceLocation.md)**|Specifies a horizontal or vertical coordinate.|
| _Index_|Optional| **[XlPieSliceIndex](Excel.XlPieSliceIndex.md)**|Specifies which pie slice position coordinate to return. The default value is **xlOuterCenterPoint**.|

## Return value

Double


## Remarks

This property only applies to pie and doughnut chart types.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]