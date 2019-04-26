---
title: FormatConditions.AddColorScale method (Excel)
keywords: vbaxl10.chm510079
f1_keywords:
- vbaxl10.chm510079
ms.prod: excel
api_name:
- Excel.FormatConditions.AddColorScale
ms.assetid: f1b23e2f-0c62-fdc5-597b-a8a444d5a4a3
ms.date: 04/26/2019
localization_priority: Normal
---


# FormatConditions.AddColorScale method (Excel)

Returns a new **[ColorScale](Excel.ColorScale.md)** object representing a conditional formatting rule that uses gradations in cell colors to indicate relative differences in the values of cells included in a selected range.


## Syntax

_expression_.**AddColorScale** (_ColorScaleType_)

_expression_ A variable that represents a **[FormatConditions](Excel.FormatConditions.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ColorScaleType_|Required| **Long**|The type of color scale.|

## Return value

**ColorScale** object


## Remarks

Color scales are visual guides that help you understand data distribution and variation. A color scale helps you identify relative differences in the values of cells in a given range by using color variation. Different colors and gradations between colors represent differences in cell values. 

For example, in a three-color scale, you can specify that cells with the highest relative data values are green, cells with intermediate values are yellow, and cells with the lowest values are red.


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]