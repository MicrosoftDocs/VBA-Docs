---
title: Top10 object (Excel)
keywords: vbaxl10.chm821072
f1_keywords:
- vbaxl10.chm821072
ms.prod: excel
api_name:
- Excel.Top10
ms.assetid: b94f4a4f-564c-d751-2b43-4b9482e048cc
ms.date: 06/08/2017
localization_priority: Normal
---


# Top10 object (Excel)

Represents a top ten visual of a conditional formatting rule. Applying a color to a range helps you see the value of a cell relative to other cells.


## Remarks

All conditional formatting objects are contained within a  **[FormatConditions](Excel.FormatConditions.md)** collection object, which is a child of a **[Range](Excel.Range(object).md)** collection. You can create a top 10 formatting rule by using either the **[Add](Excel.FormatConditions.Add.md)** or **[AddTop10](Excel.FormatConditions.AddTop10.md)** method of the **FormatConditions** collection.


## Example

The following example builds a dynamic data set and applies color to the top 10 values through conditional formatting rules.


```vb
Sub Top10CF() 
 
' Building data 
 Range("A1").Value = "Name" 
 Range("B1").Value = "Number" 
 Range("A2").Value = "Agent1" 
 Range("A2").AutoFill Destination:=Range("A2:A26"), Type:=xlFillDefault 
 Range("B2:B26").FormulaArray = "=INT(RAND()*101)" 
 Range("B2:B26").Select 
 
' Applying Conditional Formatting Top 10 
 Selection.FormatConditions.AddTop10 
 Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority 
 With Selection.FormatConditions(1) 
 .TopBottom = xlTop10Top 
 .Rank = 10 
 .Percent = False 
 End With 
 
' Applying color fill 
 With Selection.FormatConditions(1).Font 
 .Color = -16752384 
 .TintAndShade = 0 
 End With 
 With Selection.FormatConditions(1).Interior 
 .PatternColorIndex = xlAutomatic 
 .Color = 13561798 
 .TintAndShade = 0 
 End With 
MsgBox "Added Top10 Conditional Format. Press F9 to update values.", vbInformation 
 
End Sub
```


## See also



[Excel Object Model Reference](./overview/Excel/object-model.md)

