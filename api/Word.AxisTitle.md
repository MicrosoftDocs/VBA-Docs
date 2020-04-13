---
title: AxisTitle object (Word)
keywords: vbawd10.chm1499
f1_keywords:
- vbawd10.chm1499
ms.prod: word
api_name:
- Word.AxisTitle
ms.assetid: ec746a05-40df-95cc-c017-40ef150504cf
ms.date: 06/08/2017
localization_priority: Normal
---


# AxisTitle object (Word)

Represents a chart axis title.


## Remarks

Use the **[AxisTitle](Word.Axis.AxisTitle.md)** property to return an **AxisTitle** object.

The **AxisTitle** object does not exist and cannot be used unless the **[HasTitle](Word.Axis.HasTitle.md)** property for the axis is **True**.


## Example

The following example sets the caption, sets the font to Bookman 10 point, and formats the word "millions" as italic for the axis title of the value axis for the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.Axes(xlValue) 
 .HasTitle = True 
 With .AxisTitle 
 .Caption = "Revenue (millions)" 
 .Font.Name = "bookman" 
 .Font.Size = 10 
 .Characters(10, 8).Font.Italic = True 
 End With 
 End With 
 End If 
End With 

```

## Methods

- [Delete](Word.AxisTitle.Delete.md)
- [Select](Word.AxisTitle.Select.md)

## Properties

- [Application](Word.AxisTitle.Application.md)
- [Caption](Word.AxisTitle.Caption.md)
- [Characters](Word.AxisTitle.Characters.md)
- [Creator](Word.AxisTitle.Creator.md)
- [Format](Word.AxisTitle.Format.md)
- [Formula](Word.AxisTitle.Formula.md)
- [FormulaLocal](Word.AxisTitle.FormulaLocal.md)
- [FormulaR1C1](Word.AxisTitle.FormulaR1C1.md)
- [FormulaR1C1Local](Word.AxisTitle.FormulaR1C1Local.md)
- [Height](Word.AxisTitle.Height.md)
- [HorizontalAlignment](Word.AxisTitle.HorizontalAlignment.md)
- [IncludeInLayout](Word.AxisTitle.IncludeInLayout.md)
- [Left](Word.AxisTitle.Left.md)
- [Name](Word.AxisTitle.Name.md)
- [Orientation](Word.AxisTitle.Orientation.md)
- [Parent](Word.AxisTitle.Parent.md)
- [Position](Word.AxisTitle.Position.md)
- [ReadingOrder](Word.AxisTitle.ReadingOrder.md)
- [Shadow](Word.AxisTitle.Shadow.md)
- [Text](Word.AxisTitle.Text.md)
- [Top](Word.AxisTitle.Top.md)
- [VerticalAlignment](Word.AxisTitle.VerticalAlignment.md)
- [Width](Word.AxisTitle.Width.md)

## See also

- [Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]