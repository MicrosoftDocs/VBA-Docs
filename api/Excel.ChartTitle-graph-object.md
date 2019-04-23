---
title: ChartTitle object (Excel Graph)
keywords: vbagr10.chm131081
f1_keywords:
- vbagr10.chm131081
ms.prod: excel
api_name:
- Excel.ChartTitle
ms.assetid: 6eca7bbc-0158-f25e-d7c8-3f57f06ccccf
ms.date: 04/06/2019
localization_priority: Normal
---


# ChartTitle object (Excel Graph)

Represents the title of the specified chart.


## Remarks

Use the **[ChartTitle](excel.charttitle-graph-property.md)** property to return the **ChartTitle** object. 

The **ChartTitle** object doesn't exist and cannot be used unless the **[HasTitle](Excel.HasTitle.md)** property for the chart is **True**.


## Example

The following example adds a title to the chart.

```vb
With myChart 
 .HasTitle = True 
 .ChartTitle.Text = "February Sales" 
End With
```

## See also

- [Excel Graph Visual Basic Reference](overview/excel/graph-visual-basic-reference.md)
- [Excel Object Model Reference](overview/excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]