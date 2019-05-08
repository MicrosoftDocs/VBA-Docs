---
title: DisplayUnitLabel object (Excel)
keywords: vbaxl10.chm672072
f1_keywords:
- vbaxl10.chm672072
ms.prod: excel
api_name:
- Excel.DisplayUnitLabel
ms.assetid: 522dea6a-114f-3e0f-f8ae-6c2667c733dd
ms.date: 03/29/2019
localization_priority: Normal
---


# DisplayUnitLabel object (Excel)

Represents a unit label on an axis in the specified chart.


## Remarks

Unit labels are useful for charting large values—for example, in the millions or billions. You can make the chart more readable by using a single unit label instead of large numbers at each tick mark.


## Example

Use the **[DisplayUnitLabel](Excel.Axis.DisplayUnitLabel.md)** property of the **Axis** object to return the **DisplayUnitLabel** object. The following example sets the display label caption to Millions on the value axis on Chart1, and then it turns off automatic font scaling.

```vb
With Charts("Chart1").Axes(xlValue) 
 .DisplayUnit = xlMillions 
 .HasDisplayUnitLabel = True 
 With .DisplayUnitLabel 
 .Caption = "Millions" 
 .AutoScaleFont = False 
 End With 
End With
```

## Methods

- [Delete](Excel.DisplayUnitLabel.Delete.md)
- [Select](Excel.DisplayUnitLabel.Select.md)

## Properties

- [Application](Excel.DisplayUnitLabel.Application.md)
- [Caption](Excel.DisplayUnitLabel.Caption.md)
- [Characters](Excel.DisplayUnitLabel.Characters.md)
- [Creator](Excel.DisplayUnitLabel.Creator.md)
- [Format](Excel.DisplayUnitLabel.Format.md)
- [Formula](Excel.DisplayUnitLabel.Formula.md)
- [FormulaLocal](Excel.DisplayUnitLabel.FormulaLocal.md)
- [FormulaR1C1](Excel.DisplayUnitLabel.FormulaR1C1.md)
- [FormulaR1C1Local](Excel.DisplayUnitLabel.FormulaR1C1Local.md)
- [Height](Excel.DisplayUnitLabel.Height.md)
- [HorizontalAlignment](Excel.DisplayUnitLabel.HorizontalAlignment.md)
- [Left](Excel.DisplayUnitLabel.Left.md)
- [Name](Excel.DisplayUnitLabel.Name.md)
- [Orientation](Excel.DisplayUnitLabel.Orientation.md)
- [Parent](Excel.DisplayUnitLabel.Parent.md)
- [Position](Excel.DisplayUnitLabel.Position.md)
- [ReadingOrder](Excel.DisplayUnitLabel.ReadingOrder.md)
- [Shadow](Excel.DisplayUnitLabel.Shadow.md)
- [Text](Excel.DisplayUnitLabel.Text.md)
- [Top](Excel.DisplayUnitLabel.Top.md)
- [VerticalAlignment](Excel.DisplayUnitLabel.VerticalAlignment.md)
- [Width](Excel.DisplayUnitLabel.Width.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]