---
title: Chart.DepthPercent property (Excel)
keywords: vbaxl10.chm149099
f1_keywords:
- vbaxl10.chm149099
ms.prod: excel
api_name:
- Excel.Chart.DepthPercent
ms.assetid: 3b53544f-8800-c1c9-6615-c601d213daee
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.DepthPercent property (Excel)

Returns or sets the depth of a 3-D chart as a percentage of the chart width (between 20 and 2000 percent). Read/write  **Long**.


## Syntax

_expression_. `DepthPercent`

_expression_ A variable that represents a [Chart](Excel.Chart-graph-object.md) object.


## Example

This example sets the depth of Chart1 to be 50 percent of its width. The example should be run on a 3-D chart (the  **DepthPercent** property fails on 2-D charts).


```vb
Charts("Chart1").DepthPercent = 50
```


## See also


[Chart Object](Excel.Chart(object).md)

