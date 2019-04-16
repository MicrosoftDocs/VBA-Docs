---
title: Chart.ApplyLayout method (Word)
keywords: vbawd10.chm79366564
f1_keywords:
- vbawd10.chm79366564
ms.prod: word
api_name:
- Word.Chart.ApplyLayout
ms.assetid: f23d8a12-65d5-3336-4381-76bfc4b73507
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.ApplyLayout method (Word)

Applies the layouts shown in the ribbon.


## Syntax

_expression_. `ApplyLayout`( `_Layout_` , `_ChartType_` )

_expression_ A variable that represents a **[Chart](Word.Chart.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Layout_|Required| **Long**|The type of layout. The type of layout is denoted by a number from 1 to 10.|
| _ChartType_|Optional| **Variant**|An  **[XlChartType](Excel.XlChartType.md)** constant that represents the type of chart.|

## Remarks

When you use a layout on the current chart type, a number from 1 to 10 is applied to the chart type. You can also apply the layout of one chart type on another chart type. For example, you can apply the layouts that are available from a line chart to a column chart. The layout adds only chart elements that are available for that particular chart type.


## See also


[Chart Object](Word.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]