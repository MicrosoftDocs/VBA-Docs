---
title: Floor object (Excel Graph)
keywords: vbagr10.chm5207374
f1_keywords:
- vbagr10.chm5207374
ms.prod: excel
api_name:
- Excel.Floor
ms.assetid: ce76e68b-7b15-7e2c-4464-07befbf53cc5
ms.date: 04/06/2019
localization_priority: Normal
---


# Floor object (Excel Graph)

Represents the floor of the specified 3D chart.


## Remarks

Use the **[Floor](excel.floor-graph-property.md)** property to return the **Floor** object. 



## Example

The following example sets the floor color for the chart to cyan. If the chart isn't a 3D chart, this example will fail.

```vb
myChart.Floor.Interior.Color = RGB(0, 255, 255)
```

## See also

- [Excel Graph Visual Basic Reference](overview/excel/graph-visual-basic-reference.md)
- [Excel Object Model Reference](overview/excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]