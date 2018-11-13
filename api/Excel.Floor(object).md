---
title: Floor object (Excel)
keywords: vbaxl10.chm611072
f1_keywords:
- vbaxl10.chm611072
ms.prod: excel
api_name:
- Excel.Floor
ms.assetid: 74c71ca8-a0d4-f7cf-a002-5cec7a27b70d
ms.date: 06/08/2017
---


# Floor object (Excel)

Represents the floor of a 3-D chart.


## Example

Use the  **[Floor](Excel.Chart.Floor.md)** property to return the **Floor** object. The following example sets the floor color for embedded chart one to cyan. The example will fail if the chart isn't a 3-D chart.


```vb
Worksheets("sheet1").ChartObjects(1).Activate 
ActiveChart.Floor.Interior.Color = RGB(0, 255, 255)
```


## See also


[Excel Object Model Reference](overview/Excel/object-model.md)


