---
title: Shapes.AddChart2 method (Excel)
keywords: vbaxl10.chm638096
f1_keywords:
- vbaxl10.chm638096
ms.prod: excel
ms.assetid: 2d4569df-2f77-40d5-5f81-859e13e0abb7
ms.date: 05/15/2019
localization_priority: Normal
---


# Shapes.AddChart2 method (Excel)

Adds a chart to the document. Returns a **[Shape](Excel.Shape.md)** object that represents a chart and adds it to the specified collection.


## Syntax

_expression_.**AddChart2** (_Style_, _XlChartType_, _Left_, _Top_, _Width_, _Height_, _NewLayout_)

_expression_ A variable that represents a **[Shapes](Excel.Shapes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Style_|Optional|**Variant**|The chart style. Use "-1" to get the default style for the chart type specified in **[XlChartType](excel.xlcharttype.md)**. |
| _XlChartType_|Optional|**Variant**|The type of chart.|
| _Left_|Optional|**Variant**|The position, in [points](../language/glossary/vbe-glossary.md#point), of the left edge of the chart, relative to the anchor.|
| _Top_|Optional|**Variant**|The position, in points, of the top edge of the chart, relative to the anchor.|
| _Width_|Optional|**Variant**|The width, in points, of the chart.|
| _Height_|Optional|**Variant**|The height, in points, of the chart.|
| _NewLayout_|Optional|**Variant**|If **NewLayout** is **True**, the chart is inserted by using the new dynamic formatting rules (Title is on, and Legend is on only if there are multiple series).|

## Return value

**Shape**




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
