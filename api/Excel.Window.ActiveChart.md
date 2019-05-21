---
title: Window.ActiveChart property (Excel)
keywords: vbaxl10.chm356077
f1_keywords:
- vbaxl10.chm356077
ms.prod: excel
api_name:
- Excel.Window.ActiveChart
ms.assetid: 505902dd-63c3-cd11-c3cc-a82680c11768
ms.date: 05/21/2019
localization_priority: Normal
---


# Window.ActiveChart property (Excel)

Returns a **[Chart](Excel.Chart(object).md)** object that represents the active chart (either an embedded chart or a chart sheet). An embedded chart is considered active when it's either selected or activated. When no chart is active, this property returns **Nothing**.


## Syntax

_expression_.**ActiveChart**

_expression_ A variable that represents a **[Window](Excel.Window.md)** object.


## Remarks

If you don't specify an object qualifier, this property returns the active chart in the active workbook.


## Example

This example turns on the legend for the active chart.

```vb
ActiveChart.HasLegend = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]