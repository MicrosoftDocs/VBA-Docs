---
title: Interior object (Excel)
keywords: vbaxl10.chm550072
f1_keywords:
- vbaxl10.chm550072
ms.prod: excel
api_name:
- Excel.Interior
ms.assetid: 37c79831-2cac-69fd-10ee-6d5415ed338b
ms.date: 03/30/2019
localization_priority: Normal
---


# Interior object (Excel)

Represents the interior of an object.


## Example

Use the **[Interior](Excel.Range.Interior.md)** property of the **Range** object to return the **Interior** object. The following example sets the color for the interior of cell A1 to red.

```vb
Worksheets("Sheet1").Range("A1").Interior.ColorIndex = 3
```

<br/>

This example gets the value of the color of a cell in column A by using the **ColorIndex** property, and then uses that value to sort the range by color.

```vb
Sub ColorSort()
   'Set up your variables and turn off screen updating.
   Dim iCounter As Integer
   Application.ScreenUpdating = False
   
   'For each cell in column A, go through and place the color index value of the cell in column C.
   For iCounter = 2 To 55
      Cells(iCounter, 3) = _
         Cells(iCounter, 1).Interior.ColorIndex
   Next iCounter
   
   'Sort the rows based on the data in column C
   Range("C1") = "Index"
   Columns("A:C").Sort key1:=Range("C2"), _
      order1:=xlAscending, header:=xlYes
   
   'Clear out the temporary sorting value in column C, and turn screen updating back on.
   Columns(3).ClearContents
   Application.ScreenUpdating = True
End Sub
```


## Properties

- [Application](Excel.Interior.Application.md)
- [Color](Excel.Interior.Color.md)
- [ColorIndex](Excel.Interior.ColorIndex.md)
- [Creator](Excel.Interior.Creator.md)
- [Gradient](Excel.Interior.Gradient.md)
- [InvertIfNegative](Excel.Interior.InvertIfNegative.md)
- [Parent](Excel.Interior.Parent.md)
- [Pattern](Excel.Interior.Pattern.md)
- [PatternColor](Excel.Interior.PatternColor.md)
- [PatternColorIndex](Excel.Interior.PatternColorIndex.md)
- [PatternThemeColor](Excel.Interior.PatternThemeColor.md)
- [PatternTintAndShade](Excel.Interior.PatternTintAndShade.md)
- [ThemeColor](Excel.Interior.ThemeColor.md)
- [TintAndShade](Excel.Interior.TintAndShade.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
