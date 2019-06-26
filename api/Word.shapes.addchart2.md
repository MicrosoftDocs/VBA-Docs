---
title: Shapes.AddChart2 method (Word)
keywords: vbawd10.chm161415273
f1_keywords:
- vbawd10.chm161415273
ms.prod: word
ms.assetid: 54b1e65b-57ad-4824-2acf-2e1e0a22f085
ms.date: 06/08/2017
localization_priority: Normal
---


# Shapes.AddChart2 method (Word)

Adds a chart to the document. Returns a **[Shape](Word.Shape.md)** object that represents a chart and adds it to the specified collection.


## Syntax

_expression_.**AddChart2** (_Style_, _Type_, _Left_, _Top_, _Width_, _Height_, _Anchor_, _NewLayout_)

_expression_ A variable that represents a **[Shapes](Word.shapes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Style_|Optional|INT32|The chart style. Use "-1" to get the default style for the chart type specified in  **Type**.|
| _Type_|Optional|[XLCHARTTYPE](Excel.XlChartType.md)|The type of chart.|
| _Left_|Optional|**Variant**|The position, in [points](../language/glossary/vbe-glossary.md#point), of the left edge of the chart, relative to the anchor.|
| _Top_|Optional|**Variant**|The position, in points, of the top edge of the chart, relative to the anchor.|
| _Width_|Optional|**Variant**|The width, in points, of the chart.|
| _Height_|Optional|**Variant**|The height, in points, of the chart.|
| _Anchor_|Optional|**Variant**|A **Range** object that represents the text to which the canvas is bound. If _Anchor_ is specified, the anchor is positioned at the beginning of the first paragraph in the anchoring range. If this argument is omitted, the anchoring range is selected automatically and the canvas is positioned relative to the top and left edges of the page.|
| _NewLayout_|Optional|**Variant**|If _NewLayout_ is **True**, the chart will be inserted by using the new dynamic formatting rules (Title is on, and Legend is on only if there are multiple series).|

## Return value

**SHAPE**



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]