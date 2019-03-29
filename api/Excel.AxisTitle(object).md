---
title: AxisTitle object (Excel)
keywords: vbaxl10.chm564072
f1_keywords:
- vbaxl10.chm564072
ms.prod: excel
api_name:
- Excel.AxisTitle
ms.assetid: 563d3ba5-aa77-b6fc-236a-7838d75eaa53
ms.date: 03/29/2019
localization_priority: Normal
---


# AxisTitle object (Excel)

Represents a chart axis title.


## Remarks

Use the **[AxisTitle](Excel.Axis.AxisTitle.md)** property of the **Axis** object to return an **AxisTitle** object.

The **AxisTitle** object doesn't exist and cannot be used unless the **[HasTitle](Excel.Axis.HasTitle.md)** property for the axis is **True**.


## Example

The following example activates embedded chart one, sets the value axis title text, sets the font to Bookman 10 point, and formats the word millions as italic.

```vb
Worksheets("sheet1").ChartObjects(1).Activate 
With ActiveChart.Axes(xlValue) 
 .HasTitle = True 
 With .AxisTitle 
 .Caption = "Revenue (millions)" 
 .Font.Name = "bookman" 
 .Font.Size = 10 
 .Characters(10, 8).Font.Italic = True 
 End With 
End With 

```


## Methods

- [Delete](Excel.AxisTitle.Delete.md)
- [Select](Excel.AxisTitle.Select.md)

## Properties

- [Application](Excel.AxisTitle.Application.md)
- [Caption](Excel.AxisTitle.Caption.md)
- [Characters](Excel.AxisTitle.Characters.md)
- [Creator](Excel.AxisTitle.Creator.md)
- [Format](Excel.AxisTitle.Format.md)
- [Formula](Excel.AxisTitle.Formula.md)
- [FormulaLocal](Excel.AxisTitle.FormulaLocal.md)
- [FormulaR1C1](Excel.AxisTitle.FormulaR1C1.md)
- [FormulaR1C1Local](Excel.AxisTitle.FormulaR1C1Local.md)
- [Height](Excel.AxisTitle.Height.md)
- [HorizontalAlignment](Excel.AxisTitle.HorizontalAlignment.md)
- [IncludeInLayout](Excel.AxisTitle.IncludeInLayout.md)
- [Left](Excel.AxisTitle.Left.md)
- [Name](Excel.AxisTitle.Name.md)
- [Orientation](Excel.AxisTitle.Orientation.md)
- [Parent](Excel.AxisTitle.Parent.md)
- [Position](Excel.AxisTitle.Position.md)
- [ReadingOrder](Excel.AxisTitle.ReadingOrder.md)
- [Shadow](Excel.AxisTitle.Shadow.md)
- [Text](Excel.AxisTitle.Text.md)
- [Top](Excel.AxisTitle.Top.md)
- [VerticalAlignment](Excel.AxisTitle.VerticalAlignment.md)
- [Width](Excel.AxisTitle.Width.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
