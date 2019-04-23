---
title: Chart.HeightPercent property (Excel)
keywords: vbaxl10.chm149117
f1_keywords:
- vbaxl10.chm149117
ms.prod: excel
api_name:
- Excel.Chart.HeightPercent
ms.assetid: a95f2b76-57a1-4c04-9f5f-ccd7852d4ab6
ms.date: 04/16/2019
localization_priority: Normal
---


# Chart.HeightPercent property (Excel)

Returns or sets the height of a 3D chart as a percentage of the chart width (between 5 and 500 percent). Read/write **Long**.


## Syntax

_expression_.**HeightPercent**

_expression_ A variable that represents a **[Chart](Excel.Chart(object).md)** object.


## Example

This example sets the height of Chart1 to 80 percent of its width. The example should be run on a 3D chart.

```vb
Charts("Chart1").HeightPercent = 80
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]