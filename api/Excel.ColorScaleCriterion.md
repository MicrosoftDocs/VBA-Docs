---
title: ColorScaleCriterion object (Excel)
keywords: vbaxl10.chm808072
f1_keywords:
- vbaxl10.chm808072
ms.prod: excel
api_name:
- Excel.ColorScaleCriterion
ms.assetid: 8b7ffd61-b843-3995-d872-e07d35adfedc
ms.date: 03/29/2019
localization_priority: Normal
---


# ColorScaleCriterion object (Excel)

Represents the criteria for the minimum, midpoint, or maximum thresholds for a color format conditional format.


## Remarks

Each **ColorScaleCriterion** is part of the **[ColorScaleCriteria](Excel.ColorScaleCriteria.md)** collection. Use **ColorScaleCriteria** (_index_), to return an individual criterion.

You can set the **Type**, **Value**, and **FormatColor** of each threshold for the color scale.


## Example

The following code example creates a range of numbers and then applies a two-color scale conditional formatting rule to that range. The color for the minimum threshold is then assigned to red and the maximum threshold to blue by indexing into the **ColorScaleCriteria** collection to set individual criteria.

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


## Properties

- [FormatColor](Excel.ColorScaleCriterion.FormatColor.md)
- [Index](Excel.ColorScaleCriterion.Index.md)
- [Type](Excel.ColorScaleCriterion.Type.md)
- [Value](Excel.ColorScaleCriterion.Value.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]