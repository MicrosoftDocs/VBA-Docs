---
title: Legend.Position Property (Excel)
keywords: vbaxl10.chm622080
f1_keywords:
- vbaxl10.chm622080
ms.prod: excel
api_name:
- Excel.Legend.Position
ms.assetid: 6256617d-d78f-8b2e-dd27-96c71cd2a84f
ms.date: 06/08/2017
---


# Legend.Position Property (Excel)

Returns or sets a  **[XlLegendPosition](Excel.XlLegendPosition.md)** value that represents the position of the legend on the chart.


## Syntax

 _expression_. `Position`

 _expression_ A variable that represents a [Legend](Excel.Legend(Graph property).md) object.


## Example

This example moves the chart legend to the bottom of the chart.


```vb
Charts(1).Legend.Position = xlLegendPositionBottom
```


## See also


[Legend Object](Excel.Legend(object).md)

