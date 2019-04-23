---
title: ColorScale object (Excel)
keywords: vbaxl10.chm805072
f1_keywords:
- vbaxl10.chm805072
ms.prod: excel
api_name:
- Excel.ColorScale
ms.assetid: 3982b041-9178-7a45-7453-c88963501a3c
ms.date: 03/29/2019
localization_priority: Normal
---


# ColorScale object (Excel)

Represents a color scale conditional formatting rule.


## Remarks

All conditional formatting objects are contained within a **[FormatConditions](Excel.FormatConditions.md)** collection object, which is a child of a **[Range](Excel.Range(object).md)** collection. You can create a color scale formatting rule by using either the **[Add](Excel.FormatConditions.Add.md)** or **[AddColorScale](Excel.FormatConditions.AddColorScale.md)** method of the **FormatConditions** collection.

Color scales are visual guides that help you understand data distribution and variation. You can apply either a two-color or a three-color scale to a range of data, data in a table, or data in a PivotTable report. For a two-color scale conditional format, you assign the value, type, and color to the minimum and maximum thresholds of a range. A three-color scale also has a midpoint threshold.

Each of these thresholds is determined by setting the properties of the **[ColorScaleCriteria](Excel.ColorScaleCriteria.md)** object. The **ColorScaleCriteria** object, which is a child of the **ColorScale** object, is a collection of all of the **[ColorScaleCriterion](Excel.ColorScaleCriterion.md)** objects for the color scale.


## Example

The following code example creates a range of numbers and then applies a two-color scale conditional formatting rule to that range. The color for the minimum threshold is then assigned to red and the maximum threshold to blue.

```vb
Sub CreateColorScaleCF() 
 
 Dim cfColorScale As ColorScale 
 
 'Fill cells with sample data from 1 to 10 
 With ActiveSheet 
 .Range("C1") = 1 
 .Range("C2") = 2 
 .Range("C1:C2").AutoFill Destination:=Range("C1:C10") 
 End With 
 
 Range("C1:C10").Select 
 
 'Create a two-color ColorScale object for the created sample data range 
 Set cfColorScale = Selection.FormatConditions.AddColorScale(ColorScaleType:=2) 
 
 'Set the minimum threshold to red and maximum threshold to blue 
 cfColorScale.ColorScaleCriteria(1).FormatColor.Color = RGB(255, 0, 0) 
 cfColorScale.ColorScaleCriteria(2).FormatColor.Color = RGB(0, 0, 255) 
 
End Sub
```


## Methods

- [Delete](Excel.ColorScale.Delete.md)
- [ModifyAppliesToRange](Excel.ColorScale.ModifyAppliesToRange.md)
- [SetFirstPriority](Excel.ColorScale.SetFirstPriority.md)
- [SetLastPriority](Excel.ColorScale.SetLastPriority.md)

## Properties

- [Application](Excel.ColorScale.Application.md)
- [AppliesTo](Excel.ColorScale.AppliesTo.md)
- [ColorScaleCriteria](Excel.ColorScale.ColorScaleCriteria.md)
- [Creator](Excel.ColorScale.Creator.md)
- [Formula](Excel.ColorScale.Formula.md)
- [Parent](Excel.ColorScale.Parent.md)
- [Priority](Excel.ColorScale.Priority.md)
- [PTCondition](Excel.ColorScale.PTCondition.md)
- [ScopeType](Excel.ColorScale.ScopeType.md)
- [StopIfTrue](Excel.ColorScale.StopIfTrue.md)
- [Type](Excel.ColorScale.Type.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]