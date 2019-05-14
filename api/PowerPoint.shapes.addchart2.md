---
title: Shapes.AddChart2 method (PowerPoint)
ms.assetid: 07f225bc-1c0d-cca5-b6a3-9de0a018eb4c
ms.date: 06/08/2017
ms.prod: powerpoint
localization_priority: Normal
---


# Shapes.AddChart2 method (PowerPoint)

Adds a chart to the document. Returns a [Shape](PowerPoint.Shape.md) object that represents a chart and adds it to the specified collection.


## Syntax

_expression_. `AddChart2`_(Style,_ _Type,_ _Left,_ _Top,_ _Width,_ _Height,_ _NewLayout)_

_expression_ A variable that represents a **[Shapes](PowerPoint.Shapes.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Style_|Optional|**Long**|The chart style. Use "-1" to get the default style for the chart type specified in  **Type**.|
| _Type_|Optional|[XLCHARTTYPE](Excel.XlChartType.md)|The type of chart.|
| _Left_|Optional|**Single**|The position, in [points](../language/glossary/vbe-glossary.md#point), of the left edge of the chart, relative to the anchor.|
| _Top_|Optional|**Single**|The position, in [points](../language/glossary/vbe-glossary.md#point), of the top edge of the chart, relative to the anchor.|
| _Width_|Optional|**Single**|The width, in [points](../language/glossary/vbe-glossary.md#point), of the chart.|
| _Height_|Optional|**Single**|The height, in [points](../language/glossary/vbe-glossary.md#point), of the chart.|
| _NewLayout_|Optional|**Boolean**|If  _NewLayout_ is true, the chart is inserted by using the new dynamic formatting rules (Title is on, and Legend is on only if there are multiple series).|
| _Style_|Optional|INT||
| _Type_|Optional|XLCHARTTYPE||
| _Left_|Optional|FLOAT||
| _Top_|Optional|FLOAT||
| _Width_|Optional|FLOAT||
| _Height_|Optional|FLOAT||
| _NewLayout_|Optional|BOOL||

## Return value

 **SHAPE**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]