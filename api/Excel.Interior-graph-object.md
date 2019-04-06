---
title: Interior object (Excel Graph)
keywords: vbagr10.chm5207570
f1_keywords:
- vbagr10.chm5207570
ms.prod: excel
api_name:
- Excel.Interior
ms.assetid: 13a4801e-f121-2a43-cd61-cf3ac9325197
ms.date: 04/06/2019
localization_priority: Normal
---


# Interior object (Excel Graph)

Represents the interior of the specified object.


## Remarks

Use the **[Interior](excel.interior-graph-property.md)** property to return the **Interior** object. 



## Example

The following example sets the chart area color to gray and the plot area color to green.

```vb
With myChart 
 .PlotArea.Interior.Color = RGB(0, 100, 150) 
 .ChartArea.Interior.Color = RGB(50, 10, 50) 
End With
```

## See also

- [Excel Graph Visual Basic Reference](overview/excel/graph-visual-basic-reference.md)
- [Excel Object Model Reference](overview/excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]