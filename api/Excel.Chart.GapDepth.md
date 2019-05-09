---
title: Chart.GapDepth property (Excel)
keywords: vbaxl10.chm149110
f1_keywords:
- vbaxl10.chm149110
ms.prod: excel
api_name:
- Excel.Chart.GapDepth
ms.assetid: 6020490a-1343-5b79-ff7d-197f78061420
ms.date: 04/16/2019
localization_priority: Normal
---


# Chart.GapDepth property (Excel)

Returns or sets the distance between the data series in a 3D chart as a percentage of the marker width. The value of this property must be between 0 and 500. Read/write **Long**.


## Syntax

_expression_.**GapDepth**

_expression_ A variable that represents a **[Chart](Excel.Chart(object).md)** object.


## Example

This example sets the distance between the data series on Chart1 to 200 percent of the marker width. The example should be run on a 3D chart (the **GapDepth** property fails on 2D charts).

```vb
Charts("Chart1").GapDepth = 200
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]