---
title: DisplayUnitLabel object (Excel Graph)
keywords: vbagr10.chm131087
f1_keywords:
- vbagr10.chm131087
ms.prod: excel
api_name:
- Excel.DisplayUnitLabel
ms.assetid: 1d8f0340-1760-295a-2c4e-92709d1deabc
ms.date: 04/06/2019
localization_priority: Normal
---


# DisplayUnitLabel object (Excel Graph)

Represents a unit label on the value axis in the specified chart. Unit labels are useful for charting large values—for example, numbers in the millions or billions. You can make the chart more readable by using a single unit label instead of large numbers with strings of zeros next to the tick marks on the axis. This way, you need never have numbers of more than one or two digits by the tick marks.

## Remarks

Use the **[DisplayUnitLabel](Excel.DisplayUnitLabel-graph-property.md)** property to return the **DisplayUnitLabel** object. 

## Example

The following example sets the caption for the value axis in _myChart_ to Millions and turns off automatic font scaling.

```vb
With myChart.Axes(xlValue).DisplayUnitLabel 
 .Caption = "Millions" 
 .AutoScaleFont = False 
End With
```

## See also

- [Excel Graph Visual Basic Reference](overview/excel/graph-visual-basic-reference.md)
- [Excel Object Model Reference](overview/excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]