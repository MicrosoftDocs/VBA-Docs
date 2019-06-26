---
title: InlineShapes.AddChart2 method (Word)
keywords: vbawd10.chm162070638
f1_keywords:
- vbawd10.chm162070638
ms.prod: word
ms.assetid: 108899b6-24bb-cf4c-db95-066219536c19
ms.date: 06/08/2017
localization_priority: Normal
---


# InlineShapes.AddChart2 method (Word)

Adds a chart to the document. Returns an **[InlineShape](Word.InlineShape.md)** object that represents the chart and adds it to the specified collection.


## Syntax

_expression_.**AddChart2** (_Style_, _Type_, _Range_, _NewLayout_)

_expression_ A variable that represents an **[InlineShapes](Word.inlineshapes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Style_|Optional|INT32|The chart style. Use "-1" to get the default style for the chart type specified in  **Type**.|
| _Type_|Optional|[XLCHARTTYPE](Excel.XlChartType.md)|The type of chart.|
| _Range_|Optional|**Variant**|The range where the chart will be placed in the text. The chart replaces the range, unless the range is collapsed. If this argument is omitted, the chart is placed automatically.|
| _NewLayout_|Optional|**Variant**|If _NewLayout_ is true, the chart is inserted by using the new dynamic formatting rules (Title is on, and Legend is on only if there are multiple series).|
| _Type_|Optional|XLCHARTTYPE||

## Return value

**INLINESHAPE**



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]