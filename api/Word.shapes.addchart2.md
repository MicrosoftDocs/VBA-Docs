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

Adds a chart to the document. Returns a [Shape](Word.Shape.md) object that represents a chart and adds it to the specified collection.


## Syntax

_expression_. `AddChart2`_(Style,_ _Type,_ _Left,_ _Top,_ _Width,_ _Height,_ _Anchor,_ _NewLayout)_

 _expression_ A variable that represents a 'Shapes' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|||||
| _Style_|Optional|INT32|The chart style. Use "-1" to get the default style for the chart type specified in  **Type**.|
| _Type_|Optional|[XLCHARTTYPE](Excel.XlChartType.md)|The type of chart.|
| _Left_|Optional|**Variant**|The position, in [points](../language/glossary/vbe-glossary.md#point), of the left edge of the chart, relative to the anchor.|
| _Top_|Optional|**Variant**|The position, in [points](../language/glossary/vbe-glossary.md#point), of the top edge of the chart, relative to the anchor.|
| _Width_|Optional|**Variant**|The width, in [points](../language/glossary/vbe-glossary.md#point), of the chart.|
| _Height_|Optional|**Variant**|The height, in [points](../language/glossary/vbe-glossary.md#point), of the chart.|
| _Anchor_|Optional|**Variant**|If  _NewLayout_ is true, the chart will be inserted by using the new dynamic formatting rules (Title is on, and Legend is on only if there are multiple series).|
|||||
|||||
|||||
|||||
|||||
|||||
|||||
| _Type_|Optional|XLCHARTTYPE||
| _Top_|Optional|**Variant**||
| _Width_|Optional|**Variant**||
| _Height_|Optional|**Variant**||
| _Anchor_|Optional|**Variant**||
| _NewLayout_|Optional|**Variant**||

## Return value

 **SHAPE**


## See also


[Shapes Collection](Word.shapes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]