---
title: Chart.Legend property (Excel)
keywords: vbaxl10.chm149120
f1_keywords:
- vbaxl10.chm149120
ms.prod: excel
api_name:
- Excel.Chart.Legend
ms.assetid: 6396ca0f-63b5-3d4a-4f6b-b4e80a1911b3
ms.date: 04/16/2019
localization_priority: Normal
---


# Chart.Legend property (Excel)

Returns a **[Legend](Excel.Legend(object).md)** object that represents the legend for the chart. Read-only.


## Syntax

_expression_.**Legend**

_expression_ A variable that represents a **[Chart](Excel.Chart(object).md)** object.


## Example

This example turns on the legend for Chart1 and then sets the legend font color to blue.

```vb
Charts("Chart1").HasLegend = True 
Charts("Chart1").Legend.Font.ColorIndex = 5
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
