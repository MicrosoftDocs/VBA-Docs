---
title: Application.ActiveChart property (Excel)
keywords: vbaxl10.chm183075
f1_keywords:
- vbaxl10.chm183075
ms.prod: excel
api_name:
- Excel.Application.ActiveChart
ms.assetid: 37b1901c-a9c2-4a86-ce05-22f3989bc9d8
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.ActiveChart property (Excel)

Returns a **[Chart](Excel.Chart(object).md)** object that represents the active chart (either an embedded chart or a chart sheet). An embedded chart is considered active when it's either selected or activated. When no chart is active, this property returns **Nothing**.


## Syntax

_expression_.**ActiveChart**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Remarks

If you don't specify an object qualifier, this property returns the active chart in the active workbook.


## Example

This example turns on the legend for the active chart.

```vb
ActiveChart.HasLegend = True
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]