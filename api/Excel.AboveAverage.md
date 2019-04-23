---
title: AboveAverage object (Excel)
keywords: vbaxl10.chm823072
f1_keywords:
- vbaxl10.chm823072
ms.prod: excel
api_name:
- Excel.AboveAverage
ms.assetid: dd4ea82f-7986-5d6f-2b0e-fe0ca38226e2
ms.date: 03/26/2019
localization_priority: Normal
---


# AboveAverage object (Excel)

Represents an above average visual of a conditional formatting rule. Applies a color or fill to a range or selection to help you see the value of a cell relative to other cells.


## Remarks

All conditional formatting objects are contained within a **[FormatConditions](Excel.FormatConditions.md)** collection object, which is a child of a **[Range](Excel.Range(object).md)** collection. 

You can create an above average formatting rule by using either the **[Add](Excel.FormatConditions.Add.md)** or **[AddAboveAverage](Excel.FormatConditions.AddAboveAverage.md)** method of the **FormatConditions** collection.


## Example

The following example builds a dynamic data set and applies color to the above average values through conditional formatting rules.

```vb
Sub AboveAverageCF() 
 
' Building data for Melanie 
 Range("A1").Value = "Name" 
 Range("B1").Value = "Number" 
 Range("A2").Value = "Melanie-1" 
 Range("A2").AutoFill Destination:=Range("A2:A26"), Type:=xlFillDefault 
 Range("B2:B26").FormulaArray = "=INT(RAND()*101)" 
 Range("B2:B26").Select 
 
' Applying Conditional Formatting to items above the average. Should appear green fill and dark green font. 
 Selection.FormatConditions.AddAboveAverage 
 Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority 
 Selection.FormatConditions(1).AboveBelow = xlAboveAverage 
 With Selection.FormatConditions(1).Font 
 .Color = -16752384 
 .TintAndShade = 0 
 End With 
 With Selection.FormatConditions(1).Interior 
 .PatternColorIndex = xlAutomatic 
 .Color = 13561798 
 .TintAndShade = 0 
 End With 
MsgBox "Added an Above Average Conditional Format to Melanie's data. Press F9 to update values.", vbInformation 
 
End Sub
```


## Methods

- [Delete](Excel.AboveAverage.Delete.md)
- [ModifyAppliesToRange](Excel.AboveAverage.ModifyAppliesToRange.md)
- [SetFirstPriority](Excel.AboveAverage.SetFirstPriority.md)
- [SetLastPriority](Excel.AboveAverage.SetLastPriority.md)

## Properties

- [AboveBelow](Excel.AboveAverage.AboveBelow.md)
- [Application](Excel.AboveAverage.Application.md)
- [AppliesTo](Excel.AboveAverage.AppliesTo.md)
- [Borders](Excel.AboveAverage.Borders.md)
- [CalcFor](Excel.AboveAverage.CalcFor.md)
- [Creator](Excel.AboveAverage.Creator.md)
- [Font](Excel.AboveAverage.Font.md)
- [Interior](Excel.AboveAverage.Interior.md)
- [NumberFormat](Excel.AboveAverage.NumberFormat.md)
- [NumStdDev](Excel.AboveAverage.NumStdDev.md)
- [Parent](Excel.AboveAverage.Parent.md)
- [Priority](Excel.AboveAverage.Priority.md)
- [PTCondition](Excel.AboveAverage.PTCondition.md)
- [ScopeType](Excel.AboveAverage.ScopeType.md)
- [StopIfTrue](Excel.AboveAverage.StopIfTrue.md)
- [Type](Excel.AboveAverage.Type.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
