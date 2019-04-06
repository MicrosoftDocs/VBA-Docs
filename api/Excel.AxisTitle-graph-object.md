---
title: AxisTitle object (Excel Graph)
keywords: vbagr10.chm131082
f1_keywords:
- vbagr10.chm131082
ms.prod: excel
api_name:
- Excel.AxisTitle
ms.assetid: a5a62dd3-5859-6f5c-5e28-6adbf400e08e
ms.date: 04/06/2019
localization_priority: Normal
---


# AxisTitle object (Excel Graph)

Represents the title of an axis in a chart.


## Remarks

Use the **[AxisTitle](excel.axistitle-graph-property.md)** property to return an **AxisTitle** object. 

The **AxisTitle** object doesn't exist and cannot be used unless the **[HasTitle](Excel.HasTitle.md)** property for the specified axis is **True**.

## Example

The following example sets the text of the value axis title and sets the font to 10-point Bookman.

```vb
With myChart.Axes(xlValue) 
 .HasTitle = True 
 With .AxisTitle 
 .Caption = "Revenue (millions)" 
 .Font.Name = "bookman" 
 .Font.Size = 10 
 End With 
End With
```

## See also

- [Excel Graph Visual Basic Reference](overview/excel/graph-visual-basic-reference.md)
- [Excel Object Model Reference](overview/excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]